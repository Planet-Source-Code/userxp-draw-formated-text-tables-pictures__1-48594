Attribute VB_Name = "modGDIplusSmall"
Option Explicit

Public Type GdiplusStartupInput
   GdiplusVersion       As Long              ' Must be 1
   DebugEventCallback   As Long          ' Ignored on free builds
   SuppressBackgroundThread As Long    ' FALSE unless you're prepared to call
   ' the hook/unhook functions properly
   SuppressExternalCodecs As Long      ' FALSE unless you want GDI+ only to use
   ' its internal image codecs.
End Type
Public Type RECTF    ' aka RectF
   Left                 As Single
   Top                  As Single
   Right                As Single
   Bottom               As Single
End Type
Public Type RECTL     ' aka Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Public Type CLSID
   Data1                As Long
   Data2                As Integer
   Data3                As Integer
   Data4(0 To 7) As Byte
End Type
Private Type RGBQUAD
   rgbBlue              As Byte
   rgbGreen             As Byte
   rgbRed               As Byte
   rgbReserved          As Byte
End Type
Public Type BITMAPINFOHEADER '40 bytes
   biSize               As Long
   biWidth              As Long
   biHeight             As Long
   biPlanes             As Integer
   biBitCount           As Integer
   biCompression        As Long
   biSizeImage          As Long
   biXPelsPerMeter      As Long
   biYPelsPerMeter      As Long
   biClrUsed            As Long
   biClrImportant       As Long
End Type
Public Type BITMAPINFO
   bmiHeader            As BITMAPINFOHEADER
   bmiColors            As RGBQUAD
End Type
Public Type POINTF   ' aka PointF
   x                    As Single
   y                    As Single
End Type
Public Type CharacterRange
   First As Long
   length As Long
End Type



' NOTE: Enums evaluate to a Long
Public Enum GpStatus   ' aka Status
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum
Public Enum GpUnit  ' aka Unit
   UnitWorld      ' 0 -- World coordinate (non-physical unit)
   UnitDisplay    ' 1 -- Variable -- for PageTransform only
   UnitPixel      ' 2 -- Each unit is one device pixel.
   UnitPoint      ' 3 -- Each unit is a printer's point, or 1/72 inch.
   UnitInch       ' 4 -- Each unit is 1 inch.
   UnitDocument   ' 5 -- Each unit is 1/300 inch.
   UnitMillimeter ' 6 -- Each unit is 1 millimeter.
End Enum
Public Enum ImageType
   ImageTypeUnknown    ' 0
   ImageTypeBitmap     ' 1
   ImageTypeMetafile   ' 2
End Enum
Public Enum StringAlignment
   ' Left edge for left-to-right text,
   ' right for right-to-left text,
   ' and top for vertical
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum

' String format flags
'
'  DirectionRightToLeft          - For horizontal text, the reading order is
'                                  right to left. This value is called
'                                  the base embedding level by the Unicode
'                                  bidirectional engine.
'                                  For vertical text, columns are read from
'                                  right to left.
'                                  By default, horizontal or vertical text is
'                                  read from left to right.
'
'  DirectionVertical             - Individual lines of text are vertical. In
'                                  each line, characters progress from top to
'                                  bottom.
'                                  By default, lines of text are horizontal,
'                                  each new line below the previous line.
'
'  NoFitBlackBox                 - Allows parts of glyphs to overhang the
'                                  bounding rectangle.
'                                  By default glyphs are first aligned
'                                  inside the margines, then any glyphs which
'                                  still overhang the bounding box are
'                                  repositioned to avoid any overhang.
'                                  For example when an italic
'                                  lower case letter f in a font such as
'                                  Garamond is aligned at the far left of a
'                                  rectangle, the lower part of the f will
'                                  reach slightly further left than the left
'                                  edge of the rectangle. Setting this flag
'                                  will ensure the character aligns visually
'                                  with the lines above and below, but may
'                                  cause some pixels outside the formatting
'                                  rectangle to be clipped or painted.
'
'  DisplayFormatControl          - Causes control characters such as the
'                                  left-to-right mark to be shown in the
'                                  output with a representative glyph.
'
'  NoFontFallback                - Disables fallback to alternate fonts for
'                                  characters not supported in the requested
'                                  font. Any missing characters will be
'                                  be displayed with the fonts missing glyph,
'                                  usually an open square.
'
'  NoWrap                        - Disables wrapping of text between lines
'                                  when formatting within a rectangle.
'                                  NoWrap is implied when a point is passed
'                                  instead of a rectangle, or when the
'                                  specified rectangle has a zero line length.
'
'  NoClip                        - By default text is clipped to the
'                                  formatting rectangle. Setting NoClip
'                                  allows overhanging pixels to affect the
'                                  device outside the formatting rectangle.
'                                  Pixels at the end of the line may be
'                                  affected if the glyphs overhang their
'                                  cells, and either the NoFitBlackBox flag
'                                  has been set, or the glyph extends to far
'                                  to be fitted.
'                                  Pixels above/before the first line or
'                                  below/after the last line may be affected
'                                  if the glyphs extend beyond their cell
'                                  ascent / descent. This can occur rarely
'                                  with unusual diacritic mark combinations.
Public Enum StringFormatFlags
   StringFormatFlagsDirectionRightToLeft = &H1
   StringFormatFlagsDirectionVertical = &H2
   StringFormatFlagsNoFitBlackBox = &H4
   StringFormatFlagsDisplayFormatControl = &H20
   StringFormatFlagsNoFontFallback = &H400
   StringFormatFlagsMeasureTrailingSpaces = &H800
   StringFormatFlagsNoWrap = &H1000
   StringFormatFlagsLineLimit = &H2000

   StringFormatFlagsNoClip = &H4000
End Enum

Public Enum StringTrimming
   StringTrimmingNone = 0
   StringTrimmingCharacter = 1
   StringTrimmingWord = 2
   StringTrimmingEllipsisCharacter = 3
   StringTrimmingEllipsisWord = 4
   StringTrimmingEllipsisPath = 5
End Enum

Public Enum BrushType
   BrushTypeSolidColor = 0
   BrushTypeHatchFill = 1
   BrushTypeTextureFill = 2
   BrushTypePathGradient = 3
   BrushTypeLinearGradient = 4
End Enum
Public Enum DashStyle
   DashStyleSolid          ' 0
   DashStyleDash           ' 1
   DashStyleDot            ' 2
   DashStyleDashDot        ' 3
   DashStyleDashDotDot     ' 4
   DashStyleCustom         ' 5
End Enum

' FontStyle: face types and common styles
Public Enum FontStyle
   FontStyleRegular = 0
   FontStyleBold = 1
   FontStyleItalic = 2
   FontStyleBoldItalic = 3
   FontStyleUnderline = 4
   FontStyleStrikeout = 8
End Enum



Public Type PICTDESC
    Size       As Long
    Type       As Long
    hBmpOrIcon As Long
    hPal       As Long
End Type
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Public Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Public Declare Function GdipFillRegion Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal region As Long) As GpStatus

Public Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawRectangles Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, rects As RECTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawRectanglesI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, rects As RECTL, ByVal Count As Long) As GpStatus


Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Declare Function OleCreatePictureIndirect2 Lib "olepro32" Alias "OleCreatePictureIndirect" _
    (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As GpStatus

Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal Filename As String, image As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As GpStatus

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus

Public Declare Function GdipDrawImageI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Long, ByVal y As Long) As GpStatus
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus

Public Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BITMAPINFO, ByVal gdiBitmapData As Long, BITMAP As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As GpStatus
Public Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, hbmReturn As Long, ByVal background As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal hIcon As Long, BITMAP As Long) As GpStatus
Public Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, hbmReturn As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromResource Lib "gdiplus" (ByVal hInstance As Long, ByVal lpBitmapName As String, BITMAP As Long) As GpStatus

Public Declare Function GdipGetImageBounds Lib "gdiplus" (ByVal image As Long, srcRect As RECTF, srcUnit As GpUnit) As GpStatus
Public Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal image As Long, Width As Single, Height As Single) As GpStatus
Public Declare Function GdipGetImageType Lib "gdiplus" (ByVal image As Long, itype As ImageType) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As GpStatus
Public Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal image As Long, resolution As Single) As GpStatus
Public Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal image As Long, resolution As Single) As GpStatus
Public Declare Function GdipGetImageFlags Lib "gdiplus" (ByVal image As Long, flags As Long) As GpStatus
Public Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal image As Long, format As CLSID) As GpStatus
Public Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal image As Long, PixelFormat As Long) As GpStatus
Public Declare Function GdipGetImageThumbnail Lib "gdiplus" (ByVal image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, _
   Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As GpStatus

' Pen Functions (ALL)
Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Long, ByVal Width As Single, ByVal unit As GpUnit, pen As Long) As GpStatus
Public Declare Function GdipCreatePen2 Lib "gdiplus" (ByVal brush As Long, ByVal Width As Single, ByVal unit As GpUnit, pen As Long) As GpStatus
Public Declare Function GdipClonePen Lib "gdiplus" (ByVal pen As Long, clonepen As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As GpStatus
Public Declare Function GdipSetPenWidth Lib "gdiplus" (ByVal pen As Long, ByVal Width As Single) As GpStatus
Public Declare Function GdipGetPenWidth Lib "gdiplus" (ByVal pen As Long, Width As Single) As GpStatus
Public Declare Function GdipSetPenUnit Lib "gdiplus" (ByVal pen As Long, ByVal unit As GpUnit) As GpStatus
Public Declare Function GdipGetPenUnit Lib "gdiplus" (ByVal pen As Long, unit As GpUnit) As GpStatus
'Public Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal pen As Long, ByVal startCap As LineCap, ByVal endCap As LineCap, ByVal dcap As DashCap) As GpStatus
'Public Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal pen As Long, ByVal startCap As LineCap) As GpStatus
'Public Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal pen As Long, ByVal endCap As LineCap) As GpStatus
'Public Declare Function GdipSetPenDashCap Lib "gdiplus" Alias "GdipSetPenDashCap197819" (ByVal pen As Long, ByVal dcap As DashCap) As GpStatus
'Public Declare Function GdipGetPenStartCap Lib "gdiplus" (ByVal pen As Long, startCap As LineCap) As GpStatus
'Public Declare Function GdipGetPenEndCap Lib "gdiplus" (ByVal pen As Long, endCap As LineCap) As GpStatus
'Public Declare Function GdipGetPenDashCap Lib "gdiplus" Alias "GdipGetPenDashCap197819" (ByVal pen As Long, dcap As DashCap) As GpStatus
'Public Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal pen As Long, ByVal LnJoin As LineJoin) As GpStatus
'Public Declare Function GdipGetPenLineJoin Lib "gdiplus" (ByVal pen As Long, LnJoin As LineJoin) As GpStatus
'Public Declare Function GdipSetPenCustomStartCap Lib "gdiplus" (ByVal pen As Long, ByVal customCap As Long) As GpStatus
'Public Declare Function GdipGetPenCustomStartCap Lib "gdiplus" (ByVal pen As Long, customCap As Long) As GpStatus
'Public Declare Function GdipSetPenCustomEndCap Lib "gdiplus" (ByVal pen As Long, ByVal customCap As Long) As GpStatus
'Public Declare Function GdipGetPenCustomEndCap Lib "gdiplus" (ByVal pen As Long, customCap As Long) As GpStatus
'Public Declare Function GdipSetPenMiterLimit Lib "gdiplus" (ByVal pen As Long, ByVal miterLimit As Single) As GpStatus
'Public Declare Function GdipGetPenMiterLimit Lib "gdiplus" (ByVal pen As Long, miterLimit As Single) As GpStatus
'Public Declare Function GdipSetPenMode Lib "gdiplus" (ByVal pen As Long, ByVal penMode As PenAlignment) As GpStatus
'Public Declare Function GdipGetPenMode Lib "gdiplus" (ByVal pen As Long, penMode As PenAlignment) As GpStatus
'Public Declare Function GdipSetPenTransform Lib "gdiplus" (ByVal pen As Long, ByVal matrix As Long) As GpStatus
'Public Declare Function GdipGetPenTransform Lib "gdiplus" (ByVal pen As Long, ByVal matrix As Long) As GpStatus
'Public Declare Function GdipResetPenTransform Lib "gdiplus" (ByVal pen As Long) As GpStatus
'Public Declare Function GdipMultiplyPenTransform Lib "gdiplus" (ByVal pen As Long, ByVal matrix As Long, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipTranslatePenTransform Lib "gdiplus" (ByVal pen As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipScalePenTransform Lib "gdiplus" (ByVal pen As Long, ByVal sx As Single, ByVal sy As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipRotatePenTransform Lib "gdiplus" (ByVal pen As Long, ByVal angle As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipSetPenColor Lib "gdiplus" (ByVal pen As Long, ByVal argb As Long) As GpStatus
'Public Declare Function GdipGetPenColor Lib "gdiplus" (ByVal pen As Long, argb As Long) As GpStatus
'Public Declare Function GdipSetPenBrushFill Lib "gdiplus" (ByVal pen As Long, ByVal brush As Long) As GpStatus
'Public Declare Function GdipGetPenBrushFill Lib "gdiplus" (ByVal pen As Long, brush As Long) As GpStatus
'Public Declare Function GdipGetPenFillType Lib "gdiplus" (ByVal pen As Long, ptype As PenType) As GpStatus
'Public Declare Function GdipGetPenDashStyle Lib "gdiplus" (ByVal pen As Long, dStyle As DashStyle) As GpStatus
'Public Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As DashStyle) As GpStatus
'Public Declare Function GdipGetPenDashOffset Lib "gdiplus" (ByVal pen As Long, Offset As Single) As GpStatus
'Public Declare Function GdipSetPenDashOffset Lib "gdiplus" (ByVal pen As Long, ByVal Offset As Single) As GpStatus
'Public Declare Function GdipGetPenDashCount Lib "gdiplus" (ByVal pen As Long, Count As Long) As GpStatus
'Public Declare Function GdipSetPenDashArray Lib "gdiplus" (ByVal pen As Long, dash As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipGetPenDashArray Lib "gdiplus" (ByVal pen As Long, dash As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipGetPenCompoundCount Lib "gdiplus" (ByVal pen As Long, Count As Long) As GpStatus
'Public Declare Function GdipSetPenCompoundArray Lib "gdiplus" (ByVal pen As Long, dash As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipGetPenCompoundArray Lib "gdiplus" (ByVal pen As Long, dash As Single, ByVal Count As Long) As GpStatus


' Brush Functions (ALL)
Public Declare Function GdipCloneBrush Lib "gdiplus" (ByVal brush As Long, cloneBrush As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipGetBrushType Lib "gdiplus" (ByVal brush As Long, brshType As BrushType) As GpStatus

' HatchBrush Functions (ALL)
'Public Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal style As HatchStyle, ByVal forecolr As Long, ByVal backcolr As Long, brush As Long) As GpStatus
'Public Declare Function GdipGetHatchStyle Lib "gdiplus" (ByVal brush As Long, style As HatchStyle) As GpStatus
'Public Declare Function GdipGetHatchForegroundColor Lib "gdiplus" (ByVal brush As Long, forecolr As Long) As GpStatus
'Public Declare Function GdipGetHatchBackgroundColor Lib "gdiplus" (ByVal brush As Long, backcolr As Long) As GpStatus

' SolidBrush Functions (ALL)
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, brush As Long) As GpStatus
Public Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal brush As Long, ByVal argb As Long) As GpStatus
Public Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal brush As Long, argb As Long) As GpStatus

' LineBrush Functions (ALL)
'Public Declare Function GdipCreateLineBrush Lib "gdiplus" (point1 As POINTF, point2 As POINTF, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
'Public Declare Function GdipCreateLineBrushI Lib "gdiplus" (point1 As POINTL, point2 As POINTL, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
'Public Declare Function GdipCreateLineBrushFromRect Lib "gdiplus" (rect As RECTF, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As LinearGradientMode, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
'Public Declare Function GdipCreateLineBrushFromRectI Lib "gdiplus" (rect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As LinearGradientMode, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
'Public Declare Function GdipCreateLineBrushFromRectWithAngle Lib "gdiplus" (rect As RECTF, ByVal color1 As Long, ByVal color2 As Long, ByVal angle As Single, ByVal isAngleScalable As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
'Public Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "gdiplus" (rect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal angle As Single, ByVal isAngleScalable As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
'Public Declare Function GdipSetLineColors Lib "gdiplus" (ByVal brush As Long, ByVal color1 As Long, ByVal color2 As Long) As GpStatus
'Public Declare Function GdipGetLineColors Lib "gdiplus" (ByVal brush As Long, lColors As Long) As GpStatus
'Public Declare Function GdipGetLineRect Lib "gdiplus" (ByVal brush As Long, rect As RECTF) As GpStatus
'Public Declare Function GdipGetLineRectI Lib "gdiplus" (ByVal brush As Long, rect As RECTL) As GpStatus
'Public Declare Function GdipSetLineGammaCorrection Lib "gdiplus" (ByVal brush As Long, ByVal useGammaCorrection As Long) As GpStatus
'Public Declare Function GdipGetLineGammaCorrection Lib "gdiplus" (ByVal brush As Long, useGammaCorrection As Long) As GpStatus
'Public Declare Function GdipGetLineBlendCount Lib "gdiplus" (ByVal brush As Long, Count As Long) As GpStatus
'Public Declare Function GdipGetLineBlend Lib "gdiplus" (ByVal brush As Long, blend As Single, positions As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipSetLineBlend Lib "gdiplus" (ByVal brush As Long, blend As Single, positions As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipGetLinePresetBlendCount Lib "gdiplus" (ByVal brush As Long, Count As Long) As GpStatus
'Public Declare Function GdipGetLinePresetBlend Lib "gdiplus" (ByVal brush As Long, blend As Long, positions As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal brush As Long, blend As Long, positions As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipSetLineSigmaBlend Lib "gdiplus" (ByVal brush As Long, ByVal focus As Single, ByVal theScale As Single) As GpStatus
'Public Declare Function GdipSetLineLinearBlend Lib "gdiplus" (ByVal brush As Long, ByVal focus As Single, ByVal theScale As Single) As GpStatus
'Public Declare Function GdipSetLineWrapMode Lib "gdiplus" (ByVal brush As Long, ByVal WrapMd As WrapMode) As GpStatus
'Public Declare Function GdipGetLineWrapMode Lib "gdiplus" (ByVal brush As Long, WrapMd As WrapMode) As GpStatus
'Public Declare Function GdipGetLineTransform Lib "gdiplus" (ByVal brush As Long, matrix As Long) As GpStatus
'Public Declare Function GdipSetLineTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long) As GpStatus
'Public Declare Function GdipResetLineTransform Lib "gdiplus" (ByVal brush As Long) As GpStatus
'Public Declare Function GdipMultiplyLineTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipTranslateLineTransform Lib "gdiplus" (ByVal brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipScaleLineTransform Lib "gdiplus" (ByVal brush As Long, ByVal sx As Single, ByVal sy As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipRotateLineTransform Lib "gdiplus" (ByVal brush As Long, ByVal angle As Single, ByVal order As MatrixOrder) As GpStatus

' TextureBrush Functions (ALL)
'Public Declare Function GdipCreateTexture Lib "gdiplus" (ByVal image As Long, ByVal WrapMd As WrapMode, texture As Long) As GpStatus
'Public Declare Function GdipCreateTexture2 Lib "gdiplus" (ByVal image As Long, ByVal WrapMd As WrapMode, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, texture As Long) As GpStatus
'Public Declare Function GdipCreateTextureIA Lib "gdiplus" (ByVal image As Long, ByVal imageAttributes As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, texture As Long) As GpStatus
'Public Declare Function GdipCreateTexture2I Lib "gdiplus" (ByVal image As Long, ByVal WrapMd As WrapMode, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, texture As Long) As GpStatus
'Public Declare Function GdipCreateTextureIAI Lib "gdiplus" (ByVal image As Long, ByVal imageAttributes As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, texture As Long) As GpStatus
'Public Declare Function GdipGetTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long) As GpStatus
'Public Declare Function GdipSetTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long) As GpStatus
'Public Declare Function GdipResetTextureTransform Lib "gdiplus" (ByVal brush As Long) As GpStatus
'Public Declare Function GdipTranslateTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipMultiplyTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipScaleTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal sx As Single, ByVal sy As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipRotateTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal angle As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipSetTextureWrapMode Lib "gdiplus" (ByVal brush As Long, ByVal WrapMd As WrapMode) As GpStatus
'Public Declare Function GdipGetTextureWrapMode Lib "gdiplus" (ByVal brush As Long, WrapMd As WrapMode) As GpStatus
'Public Declare Function GdipGetTextureImage Lib "gdiplus" (ByVal brush As Long, image As Long) As GpStatus

' PathGradientBrush Functions (ALL)
'Public Declare Function GdipCreatePathGradient Lib "gdiplus" (Points As POINTF, ByVal Count As Long, ByVal WrapMd As WrapMode, polyGradient As Long) As GpStatus
'Public Declare Function GdipCreatePathGradientI Lib "gdiplus" (Points As POINTL, ByVal Count As Long, ByVal WrapMd As WrapMode, polyGradient As Long) As GpStatus
'Public Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal path As Long, polyGradient As Long) As GpStatus
'Public Declare Function GdipGetPathGradientCenterColor Lib "gdiplus" (ByVal brush As Long, lColors As Long) As GpStatus
'Public Declare Function GdipSetPathGradientCenterColor Lib "gdiplus" (ByVal brush As Long, ByVal lColors As Long) As GpStatus
'Public Declare Function GdipGetPathGradientSurroundColorsWithCount Lib "gdiplus" (ByVal brush As Long, argb As Long, Count As Long) As GpStatus
'Public Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "gdiplus" (ByVal brush As Long, argb As Long, Count As Long) As GpStatus
'Public Declare Function GdipGetPathGradientPath Lib "gdiplus" (ByVal brush As Long, ByVal path As Long) As GpStatus
'Public Declare Function GdipSetPathGradientPath Lib "gdiplus" (ByVal brush As Long, ByVal path As Long) As GpStatus
'Public Declare Function GdipGetPathGradientCenterPoint Lib "gdiplus" (ByVal brush As Long, Points As POINTF) As GpStatus
'Public Declare Function GdipGetPathGradientCenterPointI Lib "gdiplus" (ByVal brush As Long, Points As POINTL) As GpStatus
'Public Declare Function GdipSetPathGradientCenterPoint Lib "gdiplus" (ByVal brush As Long, Points As POINTF) As GpStatus
'Public Declare Function GdipSetPathGradientCenterPointI Lib "gdiplus" (ByVal brush As Long, Points As POINTL) As GpStatus
'Public Declare Function GdipGetPathGradientRect Lib "gdiplus" (ByVal brush As Long, rect As RECTF) As GpStatus
'Public Declare Function GdipGetPathGradientRectI Lib "gdiplus" (ByVal brush As Long, rect As RECTL) As GpStatus
'Public Declare Function GdipGetPathGradientPointCount Lib "gdiplus" (ByVal brush As Long, Count As Long) As GpStatus
'Public Declare Function GdipGetPathGradientSurroundColorCount Lib "gdiplus" (ByVal brush As Long, Count As Long) As GpStatus
'Public Declare Function GdipSetPathGradientGammaCorrection Lib "gdiplus" (ByVal brush As Long, ByVal useGammaCorrection As Long) As GpStatus
'Public Declare Function GdipGetPathGradientGammaCorrection Lib "gdiplus" (ByVal brush As Long, useGammaCorrection As Long) As GpStatus
'Public Declare Function GdipGetPathGradientBlendCount Lib "gdiplus" (ByVal brush As Long, Count As Long) As GpStatus
'Public Declare Function GdipGetPathGradientBlend Lib "gdiplus" (ByVal brush As Long, blend As Single, positions As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipSetPathGradientBlend Lib "gdiplus" (ByVal brush As Long, blend As Single, positions As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipGetPathGradientPresetBlendCount Lib "gdiplus" (ByVal brush As Long, Count As Long) As GpStatus
'Public Declare Function GdipGetPathGradientPresetBlend Lib "gdiplus" (ByVal brush As Long, blend As Long, positions As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipSetPathGradientPresetBlend Lib "gdiplus" (ByVal brush As Long, blend As Long, positions As Single, ByVal Count As Long) As GpStatus
'Public Declare Function GdipSetPathGradientSigmaBlend Lib "gdiplus" (ByVal brush As Long, ByVal focus As Single, ByVal sscale As Single) As GpStatus
'Public Declare Function GdipSetPathGradientLinearBlend Lib "gdiplus" (ByVal brush As Long, ByVal focus As Single, ByVal sscale As Single) As GpStatus
'Public Declare Function GdipGetPathGradientWrapMode Lib "gdiplus" (ByVal brush As Long, WrapMd As WrapMode) As GpStatus
'Public Declare Function GdipSetPathGradientWrapMode Lib "gdiplus" (ByVal brush As Long, ByVal WrapMd As WrapMode) As GpStatus
'Public Declare Function GdipGetPathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long) As GpStatus
'Public Declare Function GdipSetPathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long) As GpStatus
'Public Declare Function GdipResetPathGradientTransform Lib "gdiplus" (ByVal brush As Long) As GpStatus
'Public Declare Function GdipMultiplyPathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipTranslatePathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipScalePathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal sx As Single, ByVal sy As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipRotatePathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal angle As Single, ByVal order As MatrixOrder) As GpStatus
'Public Declare Function GdipGetPathGradientFocusScales Lib "gdiplus" (ByVal brush As Long, xScale As Single, yScale As Single) As GpStatus
'Public Declare Function GdipSetPathGradientFocusScales Lib "gdiplus" (ByVal brush As Long, ByVal xScale As Single, ByVal yScale As Single) As GpStatus


' FontFamily Functions (ALL)
Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As String, ByVal fontCollection As Long, fontFamily As Long) As GpStatus
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As GpStatus
Public Declare Function GdipCloneFontFamily Lib "gdiplus" (ByVal fontFamily As Long, clonedFontFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySerif Lib "gdiplus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilyMonospace Lib "gdiplus" (nativeFamily As Long) As GpStatus
' NOTE: name must be LF_FACESIZE in length or less
Public Declare Function GdipGetFamilyName Lib "gdiplus" (ByVal family As Long, ByVal name As String, ByVal language As Integer) As GpStatus
Public Declare Function GdipIsStyleAvailable Lib "gdiplus" (ByVal family As Long, ByVal style As Long, IsStyleAvailable As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerable Lib "gdiplus" (ByVal fontCollection As Long, ByVal graphics As Long, numFound As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerate Lib "gdiplus" (ByVal fontCollection As Long, ByVal numSought As Long, gpfamilies As Long, ByVal numFound As Long, ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetEmHeight Lib "gdiplus" (ByVal family As Long, ByVal style As Long, EmHeight As Integer) As GpStatus
Public Declare Function GdipGetCellAscent Lib "gdiplus" (ByVal family As Long, ByVal style As Long, CellAscent As Integer) As GpStatus
Public Declare Function GdipGetCellDescent Lib "gdiplus" (ByVal family As Long, ByVal style As Long, CellDescent As Integer) As GpStatus
Public Declare Function GdipGetLineSpacing Lib "gdiplus" (ByVal family As Long, ByVal style As Long, LineSpacing As Integer) As GpStatus

' Font Functions (ALL)
Public Declare Function GdipCreateFontFromDC Lib "gdiplus" (ByVal hdc As Long, createdfont As Long) As GpStatus
'Public Declare Function GdipCreateFontFromLogfontA Lib "gdiplus" (ByVal hdc As Long, logfont As LOGFONTA, createdfont As Long) As GpStatus
'Public Declare Function GdipCreateFontFromLogfontW Lib "gdiplus" (ByVal hdc As Long, logfont As LOGFONTW, createdfont As Long) As GpStatus
Public Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As FontStyle, ByVal unit As GpUnit, createdfont As Long) As GpStatus
Public Declare Function GdipCloneFont Lib "gdiplus" (ByVal curFont As Long, cloneFont As Long) As GpStatus
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As GpStatus
Public Declare Function GdipGetFamily Lib "gdiplus" (ByVal curFont As Long, family As Long) As GpStatus
Public Declare Function GdipGetFontStyle Lib "gdiplus" (ByVal curFont As Long, style As Long) As GpStatus
Public Declare Function GdipGetFontSize Lib "gdiplus" (ByVal curFont As Long, Size As Single) As GpStatus
Public Declare Function GdipGetFontUnit Lib "gdiplus" (ByVal curFont As Long, unit As GpUnit) As GpStatus
Public Declare Function GdipGetFontHeight Lib "gdiplus" (ByVal curFont As Long, ByVal graphics As Long, Height As Single) As GpStatus
Public Declare Function GdipGetFontHeightGivenDPI Lib "gdiplus" (ByVal curFont As Long, ByVal dpi As Single, Height As Single) As GpStatus
'Public Declare Function GdipGetLogFontA Lib "gdiplus" (ByVal curFont As Long, ByVal graphics As Long, logfont As LOGFONTA) As GpStatus
'Public Declare Function GdipGetLogFontW Lib "gdiplus" (ByVal curFont As Long, ByVal graphics As Long, logfont As LOGFONTW) As GpStatus
Public Declare Function GdipNewInstalledFontCollection Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipNewPrivateFontCollection Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipDeletePrivateFontCollection Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal fontCollection As Long, numFound As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal fontCollection As Long, ByVal numSought As Long, gpfamilies As Long, numFound As Long) As GpStatus
Public Declare Function GdipPrivateAddFontFile Lib "gdiplus" (ByVal fontCollection As Long, ByVal Filename As String) As GpStatus
Public Declare Function GdipPrivateAddMemoryFont Lib "gdiplus" (ByVal fontCollection As Long, ByVal memory As Long, ByVal length As Long) As GpStatus


' Text Functions (ALL)
Public Declare Function GdipDrawString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal TheFont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal brush As Long) As GpStatus
Public Declare Function GdipDrawStringI Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal TheFont As Long, layoutRect As RECTL, ByVal StringFormat As Long, ByVal brush As Long) As GpStatus
Public Declare Function GdipMeasureString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal TheFont As Long, layoutRect As RECTF, ByVal StringFormat As Long, boundingBox As RECTF, codepointsFitted As Long, linesFilled As Long) As GpStatus
Public Declare Function GdipMeasureCharacterRanges Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal TheFont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal regionCount As Long, regions As Long) As GpStatus
Public Declare Function GdipDrawDriverString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal TheFont As Long, ByVal brush As Long, positions As POINTF, ByVal flags As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipMeasureDriverString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal TheFont As Long, positions As POINTF, ByVal flags As Long, ByVal matrix As Long, boundingBox As RECTF) As GpStatus

' String format Functions (ALL)
Public Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericDefault Lib "gdiplus" (StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericTypographic Lib "gdiplus" (StringFormat As Long) As GpStatus
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipCloneStringFormat Lib "gdiplus" (ByVal StringFormat As Long, newFormat As Long) As GpStatus
Public Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal StringFormat As Long, ByVal flags As Long) As GpStatus
Public Declare Function GdipGetStringFormatFlags Lib "gdiplus" (ByVal StringFormat As Long, flags As Long) As GpStatus
Public Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, Align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, Align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatTrimming Lib "gdiplus" (ByVal StringFormat As Long, ByVal trimming As StringTrimming) As GpStatus
Public Declare Function GdipGetStringFormatTrimming Lib "gdiplus" (ByVal StringFormat As Long, trimming As Long) As GpStatus
'Public Declare Function GdipSetStringFormatHotkeyPrefix Lib "GdiPlus" (ByVal StringFormat As Long, ByVal hkPrefix As HotkeyPrefix) As GpStatus
'Public Declare Function GdipGetStringFormatHotkeyPrefix Lib "GdiPlus" (ByVal StringFormat As Long, hkPrefix As HotkeyPrefix) As GpStatus
Public Declare Function GdipSetStringFormatTabStops Lib "gdiplus" (ByVal StringFormat As Long, ByVal firstTabOffset As Single, ByVal Count As Long, tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStops Lib "gdiplus" (ByVal StringFormat As Long, ByVal Count As Long, firstTabOffset As Single, tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStopCount Lib "gdiplus" (ByVal StringFormat As Long, Count As Long) As GpStatus
'Public Declare Function GdipSetStringFormatDigitSubstitution Lib "GdiPlus" (ByVal StringFormat As Long, ByVal language As Integer, ByVal substitute As StringDigitSubstitute) As GpStatus
'Public Declare Function GdipGetStringFormatDigitSubstitution Lib "GdiPlus" (ByVal StringFormat As Long, language As Integer, substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatMeasurableCharacterRangeCount Lib "gdiplus" (ByVal StringFormat As Long, Count As Long) As GpStatus
Public Declare Function GdipSetStringFormatMeasurableCharacterRanges Lib "gdiplus" (ByVal StringFormat As Long, ByVal rangeCount As Long, ranges As CharacterRange) As GpStatus





Function InitGDIplus() As Long
    
    Dim m_hToken As Long, uInput As GdiplusStartupInput
    On Error Resume Next
    uInput.GdiplusVersion = 1
    Call GdiplusStartup(m_hToken, uInput)
    InitGDIplus = m_hToken

End Function

Function UnloadGDIplus(m_GDIpToken As Long)
    ' Unload the GDI+ Dll
    If m_GDIpToken <> 0 Then GdiplusShutdown m_GDIpToken
End Function

Function PaintPictureGgiPlus(FName As String, ToDC As Long, Optional x As Long, Optional y As Long, Optional TheWidth As Long, Optional TheHeight As Long, Optional IsPercent As Boolean = True, Optional MaxWidth As Long, Optional Align As AlignmentConstants = vbLeftJustify, Optional MustInitGDIplus As Boolean = True) As Boolean
Dim hImg As Long, hThumb As Long, hGraph As Long
'Dim HighQuality As Boolean, gplRet As GpStatus
    
    If MustInitGDIplus Then
        Dim hGDI As Long
        hGDI = InitGDIplus
    End If

    Call GdipLoadImageFromFile(StrConv(FName, vbUnicode), hImg)
    If hImg = 0 Then GoTo ExitFunction
        
    GdipCreateFromHDC ToDC, hGraph
    
    'Call GdipGetImageThumbnailEx(hImg, TheWidth, TheHeight, hThumb, 0, 0)
    'GdipDrawImageRectI hGraph, hThumb, 0, 0, TheWidth, TheHeight
    
    
    'If IsPercent = True and MaxWidth <> 0 then we paint the picture using
    'Width = TheWidth / 100 * MaxWidth
    'and Height =
    '1. TheHeight / 100 * RealHeight (if TheHeight <> 0)
    '2. TheWidth / 100 * RealHeight (if TheHeight = 0)
    '   in that case we auto preserve original aspect ratio
    If IsPercent And MaxWidth <> 0 And TheWidth <> 0 Then
        Dim RealHeight As Long, RealWidth As Long, Percent As Single
        Percent = TheWidth
        TheWidth = TheWidth / 100 * MaxWidth
        If TheHeight = 0 Then
            GdipGetImageWidth hImg, RealWidth
            GdipGetImageHeight hImg, RealHeight
            Percent = TheWidth / RealWidth
            TheHeight = Percent * RealHeight
        Else
            GdipGetImageHeight hImg, RealHeight
            TheHeight = TheHeight / 100 * RealHeight
        End If
    Else
        If TheWidth = 0 Then GdipGetImageWidth hImg, TheWidth
        If TheHeight = 0 Then GdipGetImageHeight hImg, TheHeight
    End If
    If MaxWidth <> 0 Then
        If Align = vbCenter Then
            x = x + (MaxWidth - TheWidth) / 2
        ElseIf Align = vbRightJustify Then
            x = x + MaxWidth - TheWidth
        End If
    End If
    
    GdipDrawImageRectI hGraph, hImg, x, y, TheWidth, TheHeight
    
    GdipDisposeImage hThumb
    GdipDisposeImage hImg
    GdipDeleteGraphics hGraph
    
    PaintPictureGgiPlus = True
           
    '-- Scale
    'HighQuality = True
    'If (HighQuality) Then
    '    gplRet = GdipSetInterpolationMode(hGraph, [InterpolationModeHighQualityBicubic])
    '  Else
    '    gplRet = GdipSetInterpolationMode(hGraph, [InterpolationModeNearestNeighbor])
    'End If
    'gplRet = GdipSetPixelOffsetMode(hGraph, [PixelOffsetModeHighQuality])
    'GdipDrawImageRectI hGraph, hImg, 0, 0, TheWidth, TheHeight
    
ExitFunction:
    If MustInitGDIplus Then
        UnloadGDIplus hGDI
    End If
    
End Function

Function PaintPictureByHandleGgiPlus(hBmp As Long, ToDC As Long, Optional x As Long, Optional y As Long, Optional TheWidth As Long, Optional TheHeight As Long, Optional IsPercent As Boolean = True, Optional MaxWidth As Long, Optional hPal As Long, Optional Align As AlignmentConstants = vbLeftJustify, Optional MustInitGDIplus As Boolean = True) As Boolean
Dim hImg As Long, hThumb As Long, hGraph As Long
'Dim HighQuality As Boolean
    
    If MustInitGDIplus Then
        Dim hGDI As Long
        hGDI = InitGDIplus
    End If
    
    Call GdipCreateBitmapFromHBITMAP(hBmp, hPal, hImg)
    If hImg = 0 Then GoTo ExitFunction
        
    GdipCreateFromHDC ToDC, hGraph
    
    'Call GdipGetImageThumbnailEx(hImg, TheWidth, TheHeight, hThumb, 0, 0)
    'GdipDrawImageRectI hGraph, hThumb, 0, 0, TheWidth, TheHeight
    
    'If IsPercent = True and MaxWidth <> 0 then we paint the picture using
    'Width = TheWidth / 100 * MaxWidth
    'and Height =
    '1. TheHeight / 100 * RealHeight (if TheHeight <> 0)
    '2. TheWidth / 100 * RealHeight (if TheHeight = 0)
    '   in that case we auto preserve original aspect ratio
    If IsPercent And MaxWidth <> 0 And TheWidth <> 0 Then
        Dim RealHeight As Long, RealWidth As Long, Percent As Single
        Percent = TheWidth
        TheWidth = TheWidth / 100 * MaxWidth
        If TheHeight = 0 Then
            GdipGetImageWidth hImg, RealWidth
            GdipGetImageHeight hImg, RealHeight
            Percent = TheWidth / RealWidth
            TheHeight = Percent * RealHeight
        Else
            GdipGetImageHeight hImg, RealHeight
            TheHeight = TheHeight / 100 * RealHeight
        End If
    Else
        If TheWidth = 0 Then GdipGetImageWidth hImg, TheWidth
        If TheHeight = 0 Then GdipGetImageHeight hImg, TheHeight
    End If
    If MaxWidth <> 0 Then
        If Align = vbCenter Then
            x = x + (MaxWidth - TheWidth) / 2
        ElseIf Align = vbRightJustify Then
            x = x + MaxWidth - TheWidth
        End If
    End If
    
    GdipDrawImageRectI hGraph, hImg, x, y, TheWidth, TheHeight
    
    GdipDisposeImage hThumb
    GdipDisposeImage hImg
    GdipDeleteGraphics hGraph
    
    PaintPictureByHandleGgiPlus = True
           
    '-- Scale
    'HighQuality = True
    'If (HighQuality) Then
    '    gplRet = GdipSetInterpolationMode(hGraph, [InterpolationModeHighQualityBicubic])
    '  Else
    '    gplRet = GdipSetInterpolationMode(hGraph, [InterpolationModeNearestNeighbor])
    'End If
    'gplRet = GdipSetPixelOffsetMode(hGraph, [PixelOffsetModeHighQuality])
    'GdipDrawImageRectI hGraph, hImg, 0, 0, TheWidth, TheHeight
    
ExitFunction:
    If MustInitGDIplus Then
        UnloadGDIplus hGDI
    End If
    
    
End Function


Public Function LoadPictureDGIplus(ByVal sFileName As String, MustInitGDIplus As Boolean) As StdPicture

  Dim gplRet        As Long
  
  Dim hImg          As Long
  Dim hBmp          As Long
  Dim uPictDesc     As PICTDESC
  Dim aGuid(0 To 3) As Long
        
    If MustInitGDIplus Then
        Dim hGDI As Long
        hGDI = InitGDIplus
    End If
        
    '-- Load image
    gplRet = GdipLoadImageFromFile(StrConv(sFileName, vbUnicode), hImg)
    
    '-- Create bitmap
    gplRet = GdipCreateHBITMAPFromBitmap(hImg, hBmp, vbBlack)
    
    '-- Free image
    gplRet = GdipDisposeImage(hImg)
    
    If (gplRet = [Ok]) Then
    
        '-- Fill struct
        With uPictDesc
            .Size = Len(uPictDesc)
            .Type = vbPicTypeBitmap
            .hBmpOrIcon = hBmp
            .hPal = 0
        End With
        
        '-- Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        aGuid(0) = &H7BF80980
        aGuid(1) = &H101ABF32
        aGuid(2) = &HAA00BB8B
        aGuid(3) = &HAB0C3000
        
        '-- Create picture from bitmap handle
        OleCreatePictureIndirect2 uPictDesc, aGuid(0), -1, LoadPictureDGIplus
    End If

    If MustInitGDIplus Then
        UnloadGDIplus hGDI
    End If

End Function

Function LoadThisPicture(ob As image, p As String)
    
    Set ob.Picture = LoadPictureDGIplus(p, True)
    If ob.Picture Is Nothing Then
      On Error Resume Next
      ob.Picture = LoadPicture(p)
      If Err <> 0 Then
          Beep
          MsgBox "Error in picture """ & p & """"
      End If
      On Error GoTo 0
    End If
    
End Function


Function MeasureCharRanges(ToDC As Long)
Dim graphics As Long, myFont As Long
Dim blueBrush As Long, redBrush As Long, blackPen As Long
Dim layoutRect_A As RECTF, layoutRect_B As RECTF, layoutRect_C As RECTF
Dim charRanges(0 To 2) As CharacterRange
Dim strFormat As Long, LenaString As Long
Dim pCharRangeRegions As String * 100
Dim region As Long, i As Long, Count As Long, aString As String


    Dim hGDI As Long
    hGDI = InitGDIplus
    
   GdipCreateFromHDC ToDC, graphics

   ' Brushes and pens used for drawing and painting
   GdipCreateSolidFill RGB(255, 0, 0), blueBrush
   GdipCreateSolidFill RGB(100, 255, 0), redBrush
   GdipCreatePen1 RGB(255, 0, 0), 1, UnitPixel, blackPen
   'SolidBrush blueBrush(Color(255, 0, 0, 255));
   'SolidBrush redBrush(Color(100, 255, 0, 0));
   'Pen        blackPen(Color(255, 0, 0, 0));

   ' Layout rectangles used for drawing strings
   layoutRect_A.Left = 20: layoutRect_A.Top = 20: layoutRect_A.Right = 130: layoutRect_A.Bottom = 130
   layoutRect_B.Left = 160: layoutRect_B.Top = 20: layoutRect_B.Right = 165: layoutRect_B.Bottom = 130
   layoutRect_C.Left = 335: layoutRect_C.Top = 20: layoutRect_C.Right = 165: layoutRect_C.Bottom = 130
   'RectF   layoutRect_A(20.0f, 20.0f, 130.0f, 130.0f);
   'RectF   layoutRect_B(160.0f, 20.0f, 165.0f, 130.0f);
   'RectF   layoutRect_C(335.0f, 20.0f, 165.0f, 130.0f);

   ' Three different ranges of character positions within the string
   charRanges(0).First = 3: charRanges(0).length = 5
   charRanges(1).First = 15: charRanges(1).length = 2
   charRanges(2).First = 30: charRanges(2).length = 15
   'CharacterRange charRanges[3] = { CharacterRange(3, 5),
   '                                 CharacterRange(15, 2),
   '                                 CharacterRange(30, 15), };

   ' Font and string format to apply to string when drawing
   Call GdipCreateFontFromDC(ToDC, myFont)
   GdipCreateStringFormat 0, 0, strFormat
   'Font         myFont(L"Times New Roman", 16.0f);
   'StringFormat strFormat;

   ' Other variables
   'Region* pCharRangeRegions; // pointer to CharacterRange regions
   'short   i;                 // loop counter
   'INT     count;             // number of character ranges set
   'WCHAR   string[] = L"The quick, brown fox easily jumps over the lazy dog.";
    aString = "The quick, brown fox easily jumps over the lazy dog."
    LenaString = -1 'Len(aString)
    aString = StrConv(aString, vbUnicode)
    
   ' Set three ranges of character positions.
   GdipSetStringFormatMeasurableCharacterRanges strFormat, 3, charRanges(0)
   'strFormat.SetMeasurableCharacterRanges(3, charRanges);

   ' Get the number of ranges that have been set, and allocate memory to
   ' store the regions that correspond to the ranges.
   GdipGetStringFormatMeasurableCharacterRangeCount strFormat, Count
   pCharRangeRegions = String(Len(pCharRangeRegions), Chr(0))
   'count = strFormat.GetMeasurableCharacterRangeCount();
   'pCharRangeRegions = new Region[count];

   ' Get the regions that correspond to the ranges within the string when
   ' layout rectangle A is used. Then draw the string, and show the regions.
   'Call GdipMeasureCharacterRanges(graphics, aString, LenaString, myFont, _
        layoutRect_A, strFormat, Count, StrPtr(pCharRangeRegions))
   If GdipDrawString(graphics, aString, LenaString, myFont, layoutRect_A, _
        strFormat, blueBrush) <> 0 Then
        Beep
        MsgBox "Not drawn"
   End If
   GoTo ExitFunction
   
   GdipDrawRectangleI graphics, blackPen, layoutRect_A.Left, _
        layoutRect_A.Top, layoutRect_A.Right, layoutRect_A.Bottom
   'graphics.MeasureCharacterRanges(string, -1,
   '   &myFont, layoutRect_A, &strFormat, count, pCharRangeRegions);
   'graphics.DrawString(string, -1,
   '   &myFont, layoutRect_A, &strFormat, &blueBrush);
   'graphics.DrawRectangle(&blackPen, layoutRect_A);
   'for ( i = 0; i < count; i++)
   '{
   '   GdipFillRegion graphics, redBrush
      'graphics.FillRegion(&redBrush, pCharRangeRegions + i);
   '}

   ' Get the regions that correspond to the ranges within the string when
   ' layout rectangle B is used. Then draw the string, and show the regions.
   GdipMeasureCharacterRanges graphics, aString, LenaString, myFont, _
        layoutRect_B, strFormat, Count, StrPtr(pCharRangeRegions)
   GdipDrawString graphics, aString, LenaString, myFont, layoutRect_B, _
        strFormat, blueBrush
   GdipDrawRectangleI graphics, blackPen, layoutRect_B.Left, _
        layoutRect_B.Top, layoutRect_B.Right, layoutRect_B.Bottom
   'graphics.MeasureCharacterRanges(string, -1,
   '   &myFont, layoutRect_B, &strFormat, count, pCharRangeRegions);
   'graphics.DrawString(string, -1,
   '   &myFont, layoutRect_B, &strFormat, &blueBrush);
   'graphics.DrawRectangle(&blackPen, layoutRect_B);
   'for ( i = 0; i < count; i++)
   '{
   '   graphics.FillRegion(&redBrush, pCharRangeRegions + i);
   '}

   ' Get the regions that correspond to the ranges within the string when
   ' layout rectangle C is used. Set trailing spaces to be included in the
   ' regions. Then draw the string, and show the regions.
   GdipSetStringFormatFlags strFormat, StringFormatFlagsMeasureTrailingSpaces
   GdipMeasureCharacterRanges graphics, aString, LenaString, myFont, _
        layoutRect_C, strFormat, Count, StrPtr(pCharRangeRegions)
   GdipDrawString graphics, aString, LenaString, myFont, layoutRect_C, _
        strFormat, blueBrush
   GdipDrawRectangleI graphics, blackPen, layoutRect_C.Left, _
        layoutRect_C.Top, layoutRect_C.Right, layoutRect_C.Bottom
   
   'strFormat.SetFormatFlags(StringFormatFlagsMeasureTrailingSpaces);
   'graphics.MeasureCharacterRanges(string, -1,
   '   &myFont, layoutRect_C, &strFormat, count, pCharRangeRegions);
   'graphics.DrawString(string, -1,
   '   &myFont, layoutRect_C, &strFormat, &blueBrush);
   'graphics.DrawRectangle(&blackPen, layoutRect_C);
   'for ( i = 0; i < count; i++)
   '{
   '   graphics.FillRegion(&redBrush, pCharRangeRegions + i);
   '}


ExitFunction:

    GdipDeleteBrush blueBrush
    GdipDeleteBrush redBrush
    GdipDeletePen blackPen
    GdipDeleteFont myFont
    GdipDeleteStringFormat strFormat
    
    GdipDeleteGraphics graphics
    
    UnloadGDIplus hGDI


End Function


