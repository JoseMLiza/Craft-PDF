VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================================================================================================
'    Component  : cCraft 1.0.0
'    Autor      : J. Elihu
'    Description: PDF Buider Class
'======================================================================================================================

Option Explicit
#Const ImplCompress = True
Public Enum PDFTextAlignment
    [T_LEFT] = 0
    [T_CENTER] = 1
    [T_RIGHT] = 2
End Enum
Public Enum PDFPageSize
  [PAGE_SIZE_A4] = 0
  [PAGE_SIZE_A3] = 1
  [PAGE_SIZE_LETTER] = 2
  [PAGE_SIZE_CUSTOM] = &HFF
End Enum
Public Enum PDFPageOrientation
  [PORTRAIT] = 0
  [LANDSCAPE] = 1
End Enum
Public Enum PDFBaseFonts
  [HELVETICA] = 0
  [HELVETICA_BOLD]
  [HELVETICA_OBLIQUE]
  [HELVETICA_BOLD_OBLIQUE]
  [TIMES_ROMAN]
  [TIMES_BOLD]
  [TIMES_ITALIC]
  [TIMES_BOLD_ITALIC]
  [COURIER]
  [COURIER_BOLD]
  [COURIER_OBLIQUE]
  [COURIER_BOLD_OBLIQUE]
  [SYMBOL]
  [ZAPF_DINGBATS]
End Enum

'-= Gdi+ =-
Private Type TBitmap
  bmType        As Long
  bmWidth       As Long
  bmHeight      As Long
  bmWidthBytes  As Long
  bmPlanes      As Integer
  bmBitsPixel   As Integer
  bmBits        As Long
End Type
Private Type BITMAP_DATA
  Width         As Long
  Height        As Long
  Stride        As Long
  PixelFormat   As Long
  Scan0         As Long
  Reserved      As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputBuf As Any, Optional ByVal outputBuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef Image As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, ByRef Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal mHicon As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal mImage As Long, ByVal mRfType As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal hImage As Long, PixelFormat As Long) As Long
Private Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal Image As Long, ByVal Guid As Any) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal mBitmap As Long, ByRef mRect As Any, ByVal mFlags As Long, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BITMAP_DATA) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BITMAP_DATA) As Long

'-= Gdi32 =-
Private Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'-= Ole32 =-
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long

'-= Kernel32 =-
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long

'-= Shlwapi =-
Private Declare Function PathIsURL Lib "shlwapi" Alias "PathIsURLW" (ByVal pszPath As Long) As Long
Private Declare Function SHCreateMemStream Lib "shlwapi" Alias "#12" (ByRef Init As Byte, ByVal cbInit As Long) As IUnknown

'-= User32 =-
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, IconInfo As Any) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function VariantChangeTypeEx Lib "oleaut32" (ByRef pvDest As Variant, ByRef pvSrc As Variant, ByVal lcid As Long, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
    
Private Const CREATOR = "Craft-PDF"
Private Const TABL = vbCrLf & "   " 'vbCrLf & vbTab

'-= Compress =-
Private Type ZlibContext
    CConv           As Integer
    CompressBound   As Long
    Compress2       As Long
End Type

Private Type tContext
    FontKey     As String
    FontSize    As Single
    '-- PORTRAIT OR LANDSCAPE
    Cx          As Single
    Cy          As Single
    '--------------------
    FontChanged As Boolean
End Type

Event BeginPage(ByVal Num As Long)
Event EndPage(ByVal Num As Long)

Private m_gdip          As Long         '-- Gdi+ Manage Images
Private m_Kids          As Variant
Private m_Objs          As Collection
Private m_oRes          As Collection
Private m_uFonts        As Collection
Private m_DocInfo       As Variant

Private m_FontWidths    As Variant
Private m_FontDefWidth  As Long

Private m_Strm          As String
Private m_iObject       As Long
Private m_Ctx           As tContext
Private m_PageCx        As Single
Private m_PageCy        As Single
Private m_Zlib          As ZlibContext

Private Sub Class_Initialize()
    Dim aInput(0 To 3) As Long: aInput(0) = 1&
    Call GdiplusStartup(m_gdip, aInput(0), ByVal 0)
    
    '-= Zlib || Zlibwapi =-
    #If ImplCompress Then
        Call zlib_init
    #End If
    
    Call StartDoc
End Sub
Private Sub Class_Terminate()
    Call GdiplusShutdown(m_gdip)
    Set m_Objs = Nothing
    Set m_oRes = Nothing
    Set m_uFonts = Nothing
End Sub
Property Get PageHeight() As Single: PageHeight = m_Ctx.Cy: End Property
Property Get PageWidth() As Single: PageWidth = m_Ctx.Cx: End Property
Property Get PageCount() As Long
On Error GoTo Q
    PageCount = UBound(m_Kids) + 1
Q:
End Property
Property Get ZlibEnabled() As Boolean
    ZlibEnabled = m_Zlib.CConv
End Property
Property Let ZlibEnabled(ByRef Value As Boolean)
    If Value Then Call zlib_init Else m_Zlib.CConv = 0
End Property

'Public Function InitZlib(ByVal LibPath As String) As Boolean
'    If m_Zlib.CC_TYPE Then Exit Function
'    Dim hMod As Long: hMod = LoadLibrary(StrPtr(LibPath))
'    If hMod Then If zlib_init(hMod) = False Then Call FreeLibrary(hMod)
'End Function

Public Sub StartDoc(Optional PageSize As PDFPageSize, Optional CustomSize As Variant, Optional Orientation As PDFPageOrientation, Optional AddNewPage As Boolean = True)

    m_DocInfo = Empty
    m_FontWidths = Empty
    m_Kids = Empty
    
    Set m_Objs = New Collection
    Set m_oRes = New Collection
    Set m_uFonts = New Collection
    
    If (PageSize = PAGE_SIZE_CUSTOM) And (IsMissing(CustomSize) Or IsEmpty(CustomSize)) Then PageSize = 0
    
    Select Case PageSize
        Case PAGE_SIZE_A3
            m_PageCx = 842: m_PageCy = 1191
        Case PAGE_SIZE_LETTER
            m_PageCx = 612: m_PageCy = 792
        Case PAGE_SIZE_CUSTOM
            If IsArray(CustomSize) Then
                m_PageCx = CustomSize(0): m_PageCy = CustomSize(1)
            Else
                m_PageCx = CustomSize: m_PageCy = CustomSize
            End If
        Case Else ' PAGE_SIZE_A4
            m_PageCx = 595: m_PageCy = 842
    End Select
    
    AddObject vbNullString '1 0 obj (Catalog)
    AddObject vbNullString '2 0 obj (Pages)
    AddObject vbNullString '3 0 obj (Resources)
    
    m_oRes.Add New Collection, "/Font"
    
    pvAddBaseFont "Helvetica"
    pvAddBaseFont "Helvetica-Bold"
    
    m_Ctx.FontKey = "/F1": m_Ctx.FontSize = 10
    If AddNewPage Then Call AddPage(Orientation)
    
End Sub
Public Sub AddPage(Optional Orientation As PDFPageOrientation)

    If PageCount Then RaiseEvent EndPage(PageCount)
    If Len(m_Strm) Then pvFlushStream m_iObject
    
    Select Case Orientation
        Case 0:    m_Ctx.Cx = m_PageCx: m_Ctx.Cy = m_PageCy 'PORTRAIT
        Case Else: m_Ctx.Cx = m_PageCy: m_Ctx.Cy = m_PageCx 'LANDSCAPE
    End Select
    
    '-- Add Page & Content
    AddObject "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " & Str(m_Ctx.Cx) & " " & Str(m_Ctx.Cy) & "] /Contents " & (ObjectCount + 2) & " 0 R >>"
    pvAddQ m_Kids, ObjectCount & " 0 R" '-- Page Ref
    m_iObject = AddObject(vbNullString) '-- Page Content
    
    m_Ctx.FontChanged = True
    RaiseEvent BeginPage(PageCount)
    
End Sub
Public Sub AddForm(ByVal Key As String, Optional ByVal xPos As Single, Optional ByVal yPos As Single, Optional ByVal BoxWidth As Single, Optional ByVal BoxHeight As Single)
On Error GoTo Q

    '--Fix: y
    If BoxWidth = 0 Then BoxWidth = m_Ctx.Cx
    If BoxHeight = 0 Then BoxHeight = m_Ctx.Cy
    If yPos < 0 Then yPos = m_Ctx.Cy + yPos - BoxHeight Else yPos = m_Ctx.Cy + (-yPos) - BoxHeight
    
    If Len(m_Strm) Then pvFlushStream m_iObject
    If Left$(Key, 1) <> "/" Then Key = "/" & Key
    
    AddObject "<< /Type /XObject /Subtype /Form /FormType 1 /BBox [" & Str(xPos) & " " & Str(yPos) & " " & Str(xPos + BoxWidth) & " " & Str(yPos + BoxHeight) & "] "
    m_iObject = ObjectCount
    
    '-- Add Refference to Resource Object
    If CollectionIndexByKey(m_oRes, "/XObject") = 0 Then m_oRes.Add New Collection, "/XObject"
    m_oRes("/XObject").Add ObjectCount & " 0 R", Key
Q:
End Sub

Public Sub AddImage(Source As Variant, ByVal Key As String)
On Error GoTo Q
Dim ImageCx     As Single
Dim ImageCy     As Single
Dim RgbData()   As Byte
Dim MaskData()  As Byte
Dim IsJpeg      As Boolean
Dim mHdr        As String
Dim oHdr        As String

    If Left$(Key, 1) <> "/" Then Key = "/" & Key
    If Not (pvLoadImage(Source, ImageCx, ImageCy, RgbData, MaskData, IsJpeg)) Then Exit Sub
    mHdr = "/Subtype /Image /Width " & Str(ImageCx) & " /Height " & Str(ImageCy) & " /BitsPerComponent 8"
    
    oHdr = mHdr & "/ColorSpace /DeviceRGB"
    If IsJpeg Then
        oHdr = oHdr & " /Filter /DCTDecode"
    ElseIf pvCompress(RgbData) Then
        oHdr = oHdr & " /Filter /FlateDecode"
    End If
    
    Dim n As Long: n = ObjectCount + 1
    oHdr = oHdr & " /Length " & (UBound(RgbData) + 1)
    If (UBound(MaskData) + 1) > 0 Then oHdr = oHdr & " /SMask " & n + 1 & " 0 R"
  
    '-- Add Image Object
    AddObject "<< /Type /XObject " & oHdr & " >>" & vbCrLf & "stream" & vbCrLf & StrConv(RgbData, vbUnicode) & vbCrLf & "endstream"
    
    '-- Add mask object if mask data exists
    If (UBound(MaskData) + 1) > 0 Then
        oHdr = mHdr & " /ColorSpace /DeviceGray"
        If pvCompress(MaskData) Then oHdr = oHdr & " /Filter /FlateDecode"
        oHdr = oHdr & " /Length " & (UBound(MaskData) + 1)
        AddObject "<< /Type /XObject " & oHdr & " >>" & vbCrLf & "stream" & vbCrLf & StrConv(MaskData, vbUnicode) & vbCrLf & "endstream"
    End If
    
    '-- Add Refference to Resource Object
    If CollectionIndexByKey(m_oRes, "/XObject") = 0 Then m_oRes.Add New Collection, "/XObject"
    m_oRes("/XObject").Add n & " 0 R", Key
Q:
End Sub
Public Sub SetFont(ByVal FontCore As PDFBaseFonts, Optional FontSize As Variant, Optional TextColor As Variant)
On Error GoTo Q

    If Not IsMissing(TextColor) Then pvAdd pvRGB(TextColor) & " rg " & vbCrLf
    Dim BF  As String
    Dim Key As String

    BF = pvBaseFont(FontCore)
    If pvSelectFont("/" & BF, Key) = False Then Key = pvAddBaseFont(BF)
    If Len(Key) = 0 Then Exit Sub
    
    If Key <> m_Ctx.FontKey Then m_FontWidths = Empty
    
    m_Ctx.FontKey = Key
    If Not IsMissing(FontSize) Then m_Ctx.FontSize = FontSize
    
    'pvAdd "BT " & m_Ctx.FontKey & " " & m_Ctx.FontSize & " Tf ET" & vbCrLf
    m_Ctx.FontChanged = True
    
Q:
End Sub
Public Sub SetTextColor(ByVal Color As Long)
    pvAdd pvRGB(Color) & " rg " & vbCrLf
End Sub
Public Sub DrawText(ByVal Text As String, ByVal x As Single, ByVal y As Single, Optional Alignment As PDFTextAlignment, Optional ByVal BoxWidth As Single)
    
    '--Fix
    If y < 0 Then y = (m_Ctx.Cy + y) - m_Ctx.FontSize Else y = (m_Ctx.Cy - y - m_Ctx.FontSize)
    If Alignment Then
        If BoxWidth = 0 Then BoxWidth = m_Ctx.Cx - x
        Select Case Alignment
            Case 1: x = x + (BoxWidth - TextWidth(Text)) / 2    'CENTER
            Case 2: x = x + (BoxWidth - TextWidth(Text))        'RIGHT
        End Select
    End If
    
    Dim Out As String
    If InStr(1, Text, "<tag") Then '-- Format Text
        Dim Words   As Variant
        Dim OpenTag As String
        
        For Each Words In pvParseTags(Text)
            If IsArray(Words) Then
                If Words(0) = 1 Then OpenTag = " q " & Words(1) & " (" 'OPEN_TAG
            Else
                If Len(OpenTag) Then
                    Out = Out & OpenTag & pvEsc(CStr(Words)) & ")Tj Q "
                    OpenTag = vbNullString
                Else
                    Out = Out & "(" & pvEsc(CStr(Words)) & ")Tj "
                End If
            End If
        Next
    Else
        Out = Out & "(" & pvEsc(Text) & ")Tj "
    End If
    
    If m_Ctx.FontChanged Then
        pvAdd "BT " & m_Ctx.FontKey & " " & Str(m_Ctx.FontSize) & " Tf " & Str(x) & " " & Str(y) & " Td " & Out & "ET" & vbCrLf
        m_Ctx.FontChanged = False
    Else
        pvAdd "BT " & Str(x) & " " & Str(y) & " Td " & Out & "ET" & vbCrLf
    End If
    
End Sub
Public Sub DrawLine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Optional Color As Long, Optional LineWidth As Single = 0.5, Optional ByVal DashPatern As String)
    
    '--Fix y pos
    If y1 < 0 Then y1 = m_Ctx.Cy + y1 Else y1 = m_Ctx.Cy - y1
    If y2 < 0 Then y2 = m_Ctx.Cy + y2 Else y2 = m_Ctx.Cy - y2
    
    If Len(DashPatern) Then
        pvAdd "q [" & DashPatern & "] 0 d " & pvRGB(Color) & " RG " & Str(LineWidth) & " w " & Str(x1) & " " & Str(y1) & " m " & Str(x2) & " " & Str(y2) & " l S Q" & vbCrLf
    Else
        pvAdd pvRGB(Color) & " RG " & Str(LineWidth) & " w " & Str(x1) & " " & Str(y1) & " m " & Str(x2) & " " & Str(y2) & " l S" & vbCrLf
    End If
    
End Sub


Public Sub DrawRectangle(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, Optional FillColor As Variant, Optional BorderColor As Variant, _
                         Optional BorderWidth As Single = 0.5, Optional ByVal DashPatern As String, Optional Radius As Single)

    '--Fix y pos
    If y < 0 Then y = m_Ctx.Cy + y - Height Else y = m_Ctx.Cy + (-y) - Height
    
    Dim Out As String
    Dim OP As String
    
    ' B : Fill and Stroke      S : Stroke       f : Fill
    If Not IsMissing(BorderColor) And Not IsMissing(FillColor) Then
        OP = "B"
    ElseIf Not IsMissing(FillColor) Then
        OP = "f"
    ElseIf Not IsMissing(BorderColor) And (BorderWidth > 0) Then
        OP = "S"
    Else
        Exit Sub
    End If
    
    If Not IsMissing(BorderColor) Then Out = Out & pvRGB(CLng(BorderColor)) & " RG "
    If Not IsMissing(FillColor) Then Out = Out & pvRGB(CLng(FillColor)) & " rg "
    
    If OP = "B" Or OP = "S" Then
        Out = BorderWidth & " w " & Out
        If Len(DashPatern) Then Out = "[" & DashPatern & "] 0 d " & Out
    End If
    
    If Radius <> 0 Then
        pvAdd "q" & vbCrLf & Out & vbCrLf & pvRoundRect(x, y, Width, Height, Radius) & vbCrLf & "h " & OP & vbCrLf & "Q" & vbCrLf
    Else
        pvAdd "q " & Out & Str(x) & " " & Str(y) & " " & Str(Width) & " " & Str(Height) & " re " & OP & " Q" & vbCrLf
    End If
    
End Sub

Public Sub DrawImage(ByVal ImageKey As String, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single)

    '-- Fix
    If y < 0 Then y = (m_Ctx.Cy + y) - Height Else y = (m_Ctx.Cy - y - Height)
    If Left$(ImageKey, 1) <> "/" Then ImageKey = "/" & ImageKey
    
    '-- Add Image to Content
    pvAdd "q " & Str(Width) & " 0 0 " & Str(Height) & " " & Str(x) & " " & Str(y) & " cm " & ImageKey & " Do Q" & vbCrLf
    
End Sub
Public Sub DrawParagraph(ByVal Text As String, ByVal x As Single, ByVal y As Single, ByVal MaxWidth As Single, Optional Justified As Boolean, Optional LineSpacing As Single, Optional OutTextHeight As Single)
    
    If IsEmpty(m_FontWidths) Then pvInitFontWidths
    
    '--Fix: y
    If y < 0 Then y = (m_Ctx.Cy + y) - m_Ctx.FontSize Else y = (m_Ctx.Cy - y - m_Ctx.FontSize)

    Dim Out As String
    If m_Ctx.FontChanged Then Out = m_Ctx.FontKey & " " & Str(m_Ctx.FontSize) & " Tf" & vbCrLf: m_Ctx.FontChanged = False
    Out = "BT" & vbCrLf & Out & Str(x) & " " & Str(y) & " Td" & vbCrLf & Str(m_Ctx.FontSize + LineSpacing) & " TL" & vbCrLf
    If InStr(1, Text, "<tag") Then
        pvAddParagraphF Text, Out, MaxWidth, Justified, LineSpacing, OutTextHeight
    Else
        pvAddParagraph Text, Out, MaxWidth, Justified, LineSpacing, OutTextHeight
    End If
    Out = Out & "ET" & vbCrLf
    Call pvAdd(Out)
End Sub
Public Sub DrawForm(ByVal FormKey As String)
    If Left$(FormKey, 1) <> "/" Then FormKey = "/" & FormKey
    pvAdd FormKey & " Do" & vbCrLf '/Frm1 Do
End Sub
Public Function TextWidth(ByVal Text As String) As Single
    
    If IsEmpty(m_FontWidths) Then pvInitFontWidths
  
    Dim Code    As Integer
    Dim Cx      As Long
    Dim i       As Long
    
    For i = 1 To Len(Text)
        Code = Asc(Mid$(Text, i, 1))
        If Code >= 0 And Code <= 255 Then
            Cx = Cx + m_FontWidths(Code - 32)
        Else
            Cx = Cx + m_FontDefWidth
        End If
    Next i
    TextWidth = (Cx / 1000#) * m_Ctx.FontSize
End Function
Public Function TextHeight(ByVal Text As String, ByVal MaxWidth As Single, Optional LineSpace As Single) As Single
    
    If IsEmpty(m_FontWidths) Then pvInitFontWidths
    Dim Words()     As String
    Dim Line        As String
    Dim currentY    As Single
    Dim TestLine    As String
    Dim LineWidth   As Single
    Dim i           As Long
    
    Words = Split(Text, " ")
    For i = 0 To UBound(Words)
        If Line = "" Then TestLine = Words(i) Else TestLine = Line & " " & Words(i)
        LineWidth = TextWidth(TestLine)
        If LineWidth <= MaxWidth Then
            Line = TestLine
        Else
            currentY = currentY + m_Ctx.FontSize + LineSpace
            Line = Words(i)
        End If
    Next
    
    '-- Add last line
    If Len(Line) Then currentY = currentY + m_Ctx.FontSize + LineSpace
    TextHeight = currentY
    
End Function
Public Sub SetInfo(Optional Title As String, Optional Autor As String, Optional Producer As String, Optional ByVal CreationDate As String)

    m_DocInfo = Empty
    If Len(Title) Then pvAddQ m_DocInfo, "/Title (" & pvEsc(Title) & ")"
    If Len(Autor) Then pvAddQ m_DocInfo, "/Author (" & pvEsc(Autor) & ")"
    If Len(Producer) Then pvAddQ m_DocInfo, "/Producer (" & pvEsc(Producer) & ")"
    If Not IsEmpty(m_DocInfo) Then pvAddQ m_DocInfo, "/Creator (" & CREATOR & ")"
    If Len(CreationDate) Then pvAddQ m_DocInfo, "/CreationDate (D:" & CreationDate & ")"
    
End Sub
Public Sub Save(ByVal FileName As String, Optional bShell As Boolean)
On Error GoTo Q
Dim hFile As Long
    
    If Left$(FileName, 2) = "\\" Then FileName = "UNC\" & Mid$(FileName, 3)
    hFile = CreateFile(StrPtr("\\?\" & FileName), &HC0000000, &H3, ByVal 0&, &H2, 0&, 0)
    If Not hFile = -1 Then
        'Dim Out() As Byte:Out = StrConv(Src, vbFromUnicode)
        Dim Out As String: Out = Contents
        If WriteFile(hFile, ByVal Out, Len(Out), 0, ByVal 0&) Then Call FlushFileBuffers(hFile)
    End If
    CloseHandle hFile
    If bShell Then Call Shell("rundll32.exe url.dll,FileProtocolHandler " & (FileName), vbMaximizedFocus)
Q:
End Sub
Property Get Contents() As String
Dim Objs()      As String
Dim Offs()      As Long
Dim Buff        As String
Dim XrefPos     As Long
Dim lCount      As Long
Dim i           As Long

    If Len(m_Strm) Then pvFlushStream m_iObject
    
    Call pvBuildObjs(Objs)
    lCount = UBound(Objs)
    ReDim Offs(1 To lCount)
    
    '-- PDF in memory --
    Buff = "%PDF-1.4" & vbCrLf
    For i = 1 To lCount
        Offs(i) = Len(Buff) '-- Reg Offset
        Buff = Buff & i & " 0 obj" & vbCrLf & Objs(i) & vbCrLf & "endobj" & vbCrLf
    Next
    
    '--   Build XREF Table ---
    XrefPos = Len(Buff)
    Buff = Buff & "xref" & vbCrLf & "0 " & (lCount + 1) & vbCrLf & "0000000000 65535 f" & vbCrLf
    For i = 1 To lCount
        Buff = Buff & Format$(Offs(i), "0000000000") & " 00000 n" & vbCrLf
    Next i
    
    '--   Trailer ----
    Buff = Buff & "trailer" & vbCrLf & _
           "<< /Size " & (lCount + 1) & " /Root 1 0 R" & _
           IIf(Not IsEmpty(m_DocInfo), " /Info " & lCount & " 0 R", "") & " >>" & vbCrLf & _
           "startxref" & vbCrLf & XrefPos & vbCrLf & "%EOF"
    
    Contents = Buff
    
End Property


'----------------------------------------------------------------------------------------------------------------------
' Todo: Private Subs
'======================================================================================================================
Private Function Str(ByRef Src As Variant) As String
    'Overwrite: Str
    'Str = LTrim$(VBA.Str(Src))
    Dim vDest As Variant
    If VarType(Src) = vbString Then
        Str = Src
    ElseIf VariantChangeTypeEx(vDest, Src, &H7F, &H12, vbString) = 0 Then
        Str = vDest '-- LOCALE_INVARIANT (&H7F)
    End If
End Function
Private Function AddObject(Content As String) As Long
    m_Objs.Add Content
    AddObject = m_Objs.Count
End Function
Private Function ObjectCount() As Long
On Error GoTo Q
    ObjectCount = m_Objs.Count
Q:
End Function
Private Function pvErr(Source As String, Desc As String)
    Debug.Print TypeName(Me) & "::" & Source & "[ERROR] > " & Desc
End Function
Private Function pvAdd(Src As String)
    m_Strm = m_Strm & Src
End Function
Private Function pvCompress(Src() As Byte) As Boolean
On Error GoTo Q

    If m_Zlib.CConv = 0 Then Exit Function
    Dim ln1 As Long
    Dim ln2 As Long
    Dim Out() As Byte
    
    ln1 = UBound(Src) - LBound(Src) + 1
    ln2 = cFunc(m_Zlib.CompressBound, m_Zlib.CConv, ln1)
    
    ReDim Out(0 To ln2 - 1)
    If cFunc(m_Zlib.Compress2, m_Zlib.CConv, VarPtr(Out(0)), VarPtr(ln2), VarPtr(Src(0)), ln1, 9&) = 0 Then
        ReDim Preserve Out(0 To ln2 - 1)
        Src = Out
        pvCompress = True
    End If
Q:
End Function
Private Function pvInvalidKey(Key As String, Prefix As String) As Boolean
    If Left$(Key, 1) <> "/" Then Key = "/" & Key
    If Left$(Key, Len(Prefix)) <> Prefix Or Val(Right$(Key, Len(Key) - Len(Prefix))) = 0 Then pvInvalidKey = True
End Function
Private Function pvSelectFont(BF As String, Key As String) As Boolean
On Error GoTo Q
Dim i As Long
    For i = 1 To m_uFonts.Count
        If m_uFonts(i) = BF Then
            Key = CollectionKey(m_uFonts, i)
            pvSelectFont = True
            Exit For
        End If
    Next
Q:
End Function
Private Function pvAddBaseFont(BF As String) As String
On Error GoTo Q

    If Len(BF) = 0 Then Exit Function
    Dim Key As String: Key = "/F" & m_uFonts.Count + 1
    
    '-- Add Core Font
    If Not (BF = "Symbol" Or BF = "ZapfDingbats") Then
        AddObject "<< /Type /Font /Subtype /Type1 /BaseFont /" & BF & " /Encoding /WinAnsiEncoding >>"
    Else
        AddObject "<< /Type /Font /Subtype /Type1 /BaseFont /" & BF & " >>"
    End If
    
    m_oRes("/Font").Add ObjectCount & " 0 R", Key
    m_uFonts.Add "/" & BF, Key
    pvAddBaseFont = Key
    
Q:
End Function
Private Sub pvFlushStream(Index As Long)
Dim sObj    As String
Dim Cmps    As Boolean

    If Index > 0 And Index <= m_Objs.Count Then
        sObj = m_Objs(Index)
        m_Objs.Remove Index
    End If
    
    If Len(sObj) = 0 Then sObj = "<< "
    
    '-- Compress --
    If m_Zlib.CConv Then
        Dim Src() As Byte
        Src = StrConv(m_Strm, vbFromUnicode)
        Cmps = pvCompress(Src)
        If Cmps Then
            sObj = sObj & "/Length " & (UBound(Src) + 1) & " /Filter /FlateDecode >>" & vbCrLf
            sObj = sObj & "stream" & vbCrLf & StrConv(Src, vbUnicode) & vbCrLf & "endstream"
        End If
    End If
    
    If Not Cmps Then
        sObj = sObj & "/Length " & Len(m_Strm) & " >>" & vbCrLf
        sObj = sObj & "stream" & vbCrLf & m_Strm & "endstream"
    End If
    
    If Index > 0 And Index <= m_Objs.Count Then
        m_Objs.Add sObj, Before:=Index
    Else
        m_Objs.Add sObj
    End If
    
    m_Strm = vbNullString
    
End Sub

Private Sub pvBuildObjs(Objs() As String)
Dim i   As Long
Dim n   As Long

    n = m_Objs.Count
    If Not IsEmpty(m_DocInfo) Then n = n + 1
    ReDim Objs(1 To n)
    For i = 4 To m_Objs.Count
        Objs(i) = m_Objs(i)
        If LenB(Objs(i)) = 0 Then Objs(i) = "<< /Length 0>>"
    Next
    
    Objs(1) = "<< /Type /Catalog /Pages 2 0 R >>"
    Objs(2) = "<< /Type /Pages " & TABL & "/Count " & (UBound(m_Kids) + 1) & TABL & "/Kids [" & Join(m_Kids, " ") & "]" & TABL & "/Resources 3 0 R" & vbCrLf & ">>"
    Objs(3) = pvBuildObj(m_oRes)
    
    If Not IsEmpty(m_DocInfo) Then
        Objs(i) = "<<" & vbCrLf & Join(m_DocInfo, vbCrLf) & vbCrLf & ">>"
    End If
    
End Sub
Private Function pvBuildObj(Src As Collection, Optional Level As Long, Optional ByVal oName As String) As String
Dim This As Variant
Dim Keys As Variant
Dim Out  As String
Dim i    As Long

    Keys = CollectionAllKeys(Src)
    Out = vbCrLf
    For i = 1 To Src.Count
        If IsObject(Src(i)) Then
            Out = Out & Keys(i) & " " & pvBuildObj(Src(i), Level + 1, Keys(i)) & vbCrLf
        Else
            Out = Out & String(Level, vbTab)
            Out = Out & Keys(i) & " " & Src(i)
            Out = Out & vbCrLf
        End If
    Next
    pvBuildObj = "<<" & Out & ">>"
End Function
Private Function pvAddQ(Src As Variant, Value As Variant)
    If IsEmpty(Src) Then Src = Array()
    ReDim Preserve Src(UBound(Src) + 1)
    Src(UBound(Src)) = Value
End Function
Private Function pvEsc(Str As String) As String
  pvEsc = Replace(Replace(Replace(Str, "\", "\\"), "(", "\("), ")", "\)")
End Function
Private Function pvReadFile(ByVal FileName As String) As Byte()
Dim hFile As Long
Dim lSize As Long
Dim Out() As Byte
    
    '// GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING
    hFile = CreateFile(StrPtr(FileName), &H80000000, &H1, ByVal 0&, &H3, &H80, 0)
    If hFile = -1 Then pvReadFile = vbNullString: Exit Function
    
    lSize = GetFileSize(hFile, 0&)
    If lSize > 0& Then
        ReDim Out(lSize - 1)
        If ReadFile(hFile, Out(0), lSize, lSize, ByVal 0&) = 0 Then
            If lSize Then ReDim Preserve Out(lSize - 1)
        End If
    End If
    CloseHandle hFile
    pvReadFile = Out
End Function
Private Function pvRGB(ByVal Color As Long) As String
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    Dim R As Byte, G As Byte, B As Byte
    B = ((Color \ &H10000) And &HFF)
    G = ((Color \ &H100) And &HFF)
    R = (Color And &HFF)
    pvRGB = Str(Round(R / 255, 3)) & " " & Str(Round(G / 255, 3)) & " " & Str(Round(B / 255, 3))
End Function
Private Function pvJustify(Src As String, MaxCx As Single) As String
Dim TxtCx       As Single
Dim nSpc        As Long
Dim Cx_p_Spc    As Single
Dim Words()     As String
Dim Tmp         As String
Dim i           As Long

    TxtCx = pvTextCx(Src, nSpc)
    Cx_p_Spc = ((MaxCx - TxtCx) / nSpc) * (1000 / m_Ctx.FontSize) * -1
    
    Words = Split(Src, " ")
    For i = 0 To UBound(Words)
        If i = 0 Then
            Tmp = Tmp & "(" & pvEsc(Words(i)) & ") "
        Else
            Tmp = Tmp & Str(Cx_p_Spc) & "(" & pvEsc(Words(i)) & ") "
        End If
    Next
    pvJustify = "[" & Tmp & "]TJ"
End Function
Private Function pvTextCx(ByVal Text As String, Optional nSpc As Long) As Double
Dim Code As Integer
Dim Cx   As Long
Dim i    As Long
    
    nSpc = 0
    For i = 1 To Len(Text)
        Code = Asc(Mid$(Text, i, 1))
        If Code >= 0 And Code <= 255 Then
            If Code = 32 Then
                nSpc = nSpc + 1
            Else
                Cx = Cx + m_FontWidths(Code - 32)
            End If
        Else
            Cx = Cx + m_FontDefWidth
        End If
    Next i
    pvTextCx = (Cx / 1000#) * m_Ctx.FontSize
End Function

Private Function CollectionAllKeys(oCol As Collection) As String()
    Const SIGN_BIT      As Long = &H80000000
    Dim aRetVal()       As String
    Dim lPtr            As Long
    Dim lIdx            As Long
    Dim sTemp           As String
    
    If oCol.Count = 0 Then
        aRetVal = Split(vbNullString)
    Else
        ReDim aRetVal(1 To oCol.Count) As String
        lPtr = ObjPtr(oCol)
        For lIdx = LBound(aRetVal) To UBound(aRetVal)
            Call CopyMemory(lPtr, ByVal (lPtr Xor SIGN_BIT) + &H18 Xor SIGN_BIT, 4)
            Call CopyMemory(ByVal VarPtr(sTemp), ByVal (lPtr Xor SIGN_BIT) + &H10 Xor SIGN_BIT, 4)
            aRetVal(lIdx) = sTemp
        Next
        Call CopyMemory(ByVal VarPtr(sTemp), 0&, 4)
    End If
    CollectionAllKeys = aRetVal
End Function
Private Function CollectionKey(oCol As Collection, ByVal lIdx As Long) As String
    Const SIGN_BIT      As Long = &H80000000
    Dim lPtr            As Long
    Dim sTemp           As String
    
    If lIdx >= 1 And lIdx <= oCol.Count Then
        lPtr = ObjPtr(oCol)
        For lIdx = 1 To lIdx
            Call CopyMemory(lPtr, ByVal (lPtr Xor SIGN_BIT) + &H18 Xor SIGN_BIT, 4)
        Next
        Call CopyMemory(ByVal VarPtr(sTemp), ByVal (lPtr Xor SIGN_BIT) + &H10 Xor SIGN_BIT, 4)
        CollectionKey = sTemp
        Call CopyMemory(ByVal VarPtr(sTemp), 0&, 4)
    End If
End Function
Private Function CollectionIndexByKey(oCol As VBA.Collection, ByVal sKey As String, Optional ByVal IgnoreCase As Boolean = True) As Long
    Dim lItemPtr        As Long
    Dim lEofPtr         As Long
    Dim lPtr            As Long
    Dim sTemp           As String
    Dim eMethod         As VbCompareMethod
    Const SIGN_BIT      As Long = &H80000000
    
    If Not oCol Is Nothing Then
        Call CopyMemory(lItemPtr, ByVal (ObjPtr(oCol) Xor SIGN_BIT) + &H24 Xor SIGN_BIT, 4)
        Call CopyMemory(lEofPtr, ByVal (ObjPtr(oCol) Xor SIGN_BIT) + &H28 Xor SIGN_BIT, 4)
    End If
    
    eMethod = IIf(IgnoreCase, vbTextCompare, vbBinaryCompare)
    Do While lItemPtr <> lEofPtr
        Call CopyMemory(ByVal VarPtr(sTemp), ByVal (lItemPtr Xor SIGN_BIT) + &H10 Xor SIGN_BIT, 4)
        Select Case StrComp(sKey, sTemp, eMethod)
        Case Is < 0: Call CopyMemory(lItemPtr, ByVal (lItemPtr Xor SIGN_BIT) + &H28 Xor SIGN_BIT, 4)
        Case Is > 0: Call CopyMemory(lItemPtr, ByVal (lItemPtr Xor SIGN_BIT) + &H24 Xor SIGN_BIT, 4)
        Case Else
            lPtr = ObjPtr(oCol)
            Do While lPtr <> lItemPtr
                Call CopyMemory(lPtr, ByVal (lPtr Xor SIGN_BIT) + &H18 Xor SIGN_BIT, 4)
                CollectionIndexByKey = CollectionIndexByKey + 1
            Loop
            GoTo QH
        End Select
    Loop
QH:
    Call CopyMemory(ByVal VarPtr(sTemp), 0&, 4)
End Function



'----------------------------------------------------------------------------------------------------------------------
' Todo: Image Methods
'======================================================================================================================
Private Function pvLoadImage(Source As Variant, Cx As Single, Cy As Single, OutRgb() As Byte, OutMask() As Byte, jpeg As Boolean) As Boolean
On Error GoTo Q
Dim Img     As Long
Dim Src()   As Byte
Dim Strm    As IUnknown

    Select Case VarType(Source)
        Case vbObject:
        Case vbString:
            Call GdipLoadImageFromFile(StrPtr(Source), Img)
        Case vbArray + vbByte
            If LenB(Source) Then Src() = Source Else Exit Function
            Set Strm = SHCreateMemStream(Src(LBound(Src)), UBound(Src) - LBound(Src) + 1)
            Call GdipLoadImageFromStream(Strm, Img)
        Case vbLong
            If GetObjectType(Source) = 7 Then ' /* OBJ_BITMAP */
                If pvBmpToImage(Source, Img) = False Then Call GdipCreateBitmapFromHBITMAP(Source, 0, Img)
            Else
                If TypeName(Source) = "Long" Then
                    If pvIconToImage(Source, Img) = False Then Call GdipCreateBitmapFromHICON(Source, Img)
                Else
                    Call GdipCreateBitmapFromHICON(Source, Img)
                End If
            End If
    End Select
    If Img = 0 Then Exit Function
    
    GdipGetImageDimension Img, Cx, Cy
    Select Case pvImageGuid(Img)
        Case "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}" '-- Image Format JPEG
            Select Case VarType(Source)
                Case vbString
                    OutRgb = pvReadFile(Source)
                Case vbArray + vbByte
                    OutRgb = Src
            End Select
            jpeg = True
            pvLoadImage = (UBound(OutRgb) + 1 > 0) 'Len(OutRgb)
            OutMask = vbNullString
        Case Else
            pvLoadImage = pvGetImageBits(Img, Cx, Cy, OutRgb, OutMask)
    End Select
    If Img Then GdipDisposeImage Img
Q:
End Function
Private Function pvImageGuid(Img As Long) As String
Dim CGuid(1) As Currency
Dim n        As Long
    If GdipGetImageRawFormat(Img, VarPtr(CGuid(0))) = 0& Then
        pvImageGuid = String$(40, vbNullChar)
        n = StringFromGUID2(VarPtr(CGuid(0)), StrPtr(pvImageGuid), 40&)
        pvImageGuid = Left$(pvImageGuid, n - 1&)
    End If
End Function

Private Function pvBmpToImage(Source As Variant, Image As Long) As Boolean
Dim tBmp As TBitmap
Dim hBmp As Long
    
    '===========================================
    ' Pixel Format
    '-------------------------------------------
    ' PixelFormat32bppPARGB = &HE200B
    ' PixelFormat32bppARGB = &H26200A
    ' PixelFormat32bppRGB = &H22009
    '===========================================
    
    If GetObjectA(Source, Len(tBmp), tBmp) Then
        If (tBmp.bmBitsPixel = 32) And (Not tBmp.bmBits = 0) Then
            GdipCreateBitmapFromScan0 tBmp.bmWidth, tBmp.bmHeight, tBmp.bmWidthBytes, &H26200A, tBmp.bmBits, Image
        ElseIf (tBmp.bmBitsPixel = 32) And (tBmp.bmBits = 0) Then
            hBmp = CopyImage(Source, 0, 0, 0, &H2000 Or &H2)
            If GetObjectA(hBmp, Len(tBmp), tBmp) Then
                If (Not tBmp.bmBits = 0) Then
                    GdipCreateBitmapFromScan0 tBmp.bmWidth, tBmp.bmHeight, tBmp.bmWidthBytes, &H26200A, tBmp.bmBits, Image
                End If
            End If
        End If
    End If
    
    If Image Then GdipImageRotateFlip Image, 6&
    If hBmp Then DeleteObject hBmp
    pvBmpToImage = Image <> 0
    
End Function
Private Function pvIconToImage(ByVal Handle As Long, Out As Long) As Boolean
Dim uInfo(4) As Long '-= fIcon|xHotspot|yHotspot|hbmMask|hbmColor =-
    If GetIconInfo(Handle, uInfo(0)) Then pvIconToImage = pvBmpToImage(uInfo(4), Out)
    If uInfo(3) Then DeleteObject uInfo(3)
    If uInfo(4) Then DeleteObject uInfo(4)
End Function

Private Function pvGetImageBits(Img As Long, Cx As Single, Cy As Single, OutRgb() As Byte, OutAlpha() As Byte) As Boolean
Dim Bmpd    As BITMAP_DATA
Dim Src()   As Byte
Dim PxFmt   As Long
Dim Rgb()   As Byte
Dim Alp()   As Byte
Dim a       As Boolean
Dim i       As Long

    If Not GdipGetImagePixelFormat(Img, PxFmt) = 0 Then Exit Function
    ReDim Src(0 To (Cx * Cy) * 4)
    
    Bmpd.Scan0 = VarPtr(Src(0))
    Bmpd.Stride = 4& * Cx
    If GdipBitmapLockBits(Img, ByVal 0&, &H4 Or &H1, &H22009, Bmpd) = 0 Then

        ReDim Rgb(0 To Cx * Cy * 3 - 1) '-- 3bpp
        ReDim Alp(0 To Cx * Cy - 1)
        
        Dim p As Long
  
        For i = 0 To UBound(Src) - 1 Step 4
            p = i \ 4
            '-- BGRA TO RGB
            Rgb(p * 3) = Src(i + 2)       ' Red
            Rgb(p * 3 + 1) = Src(i + 1)   ' Green
            Rgb(p * 3 + 2) = Src(i)       ' Blue
            
            Alp(p) = Src(i + 3)           ' Alpha
            If a = False Then If Src(i + 3) < 255 Then a = True
        Next
        pvGetImageBits = True
        Call GdipBitmapUnlockBits(Img, Bmpd)
        OutRgb = Rgb
        If a Then OutAlpha = Alp Else OutAlpha = vbNullString
    End If
End Function

Private Function pvBaseFont(i As PDFBaseFonts) As String
    pvBaseFont = Choose(i + 1, "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique", _
    "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic", _
    "Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique", "Symbol", "ZapfDingbats")
End Function
Private Function pvRoundRect(x As Single, y As Single, W As Single, H As Single, R As Single) As String
    If (R > W / 2) Or (R > H / 2) Then Exit Function
    
    Dim s As String
    Const Kappa = 0.5522847498
    s = s & x & " " & Str(Round(y + R, 2)) & " m" & vbCrLf                           ' Punto inicial (x, y+r)
    s = s & x & " " & Str(Round(y + H - R, 2)) & " l" & vbCrLf                       ' Lado izquierdo
    s = s & x & " " & Str(Round(y + H - R * (1 - Kappa), 2)) & " " & _
            Str(Round(x + R * Kappa, 2)) & " " & Str(Round(y + H, 2)) & " " & _
            Str(Round(x + R, 2)) & " " & Str(Round(y + H, 2)) & " c" & vbCrLf             ' Esquina superior izquierda
    s = s & Str(Round(x + W - R, 2)) & " " & Str(Round(y + H, 2)) & " l" & vbCrLf         ' Lado superior
    s = s & Str(Round(x + W - R * (1 - Kappa), 2)) & " " & Str(Round(y + H, 2)) & " " & _
            Str(Round(x + W, 2)) & " " & Str(Round(y + H - R * Kappa, 2)) & " " & _
            Str(Round(x + W, 2)) & " " & Str(Round(y + H - R, 2)) & " c" & vbCrLf         ' Esquina superior derecha
    s = s & Str(Round(x + W, 2)) & " " & Str(Round(y + R, 2)) & " l" & vbCrLf             ' Lado derecho
    s = s & Str(Round(x + W, 2)) & " " & Str(Round(y + R * (1 - Kappa), 2)) & " " & _
            Str(Round(x + W - R * Kappa, 2)) & " " & y & " " & _
            Str(Round(x + W - R, 2)) & " " & y & " c" & vbCrLf                       ' Esquina inferior derecha
    s = s & Str(Round(x + R, 2)) & " " & y & " l" & vbCrLf                           ' Lado inferior
    s = s & Str(Round(x + R * (1 - Kappa), 2)) & " " & y & " " & _
            x & " " & Str(Round(y + R * Kappa, 2)) & " " & _
            x & " " & Str(Round(y + R, 2)) & " c"                                    ' Esquina inferior izquierda
    pvRoundRect = s
End Function
Private Function pvParseTags(ByVal Text As String) As Collection
Dim oRegEx  As Object
Dim oExec   As Object
Dim This    As Variant
Dim Pos     As Long
Dim lastPos As Long
Dim Tmp     As String
Dim Out     As Collection
Dim v1      As Variant
Dim v2      As Variant


    Set Out = New Collection
    Set oRegEx = CreateObject("VBScript.RegExp")
    With oRegEx
        .Global = True
        .IgnoreCase = True
        .Pattern = "<tag([^>]*)>(.*?)</tag>"
        Set oExec = .Execute(Text)
    End With
    
    lastPos = 1
    For Each This In oExec
        Pos = This.FirstIndex + 1
        
        If Pos > lastPos Then
            Tmp = Mid$(Text, lastPos, Pos - lastPos)
            Out.Add Tmp
        End If
        
        Tmp = This.SubMatches(1)        ' Extract inner text
        
        ReDim v2(0 To 0) '-- ForeColor
        
        '-- Parse <tag ...>
        For Each v1 In Split(Trim(This.SubMatches(0)), " ")
            v1 = Split(v1, "=")
            Select Case Trim(v1(0))
                Case "c": 'v2(0) = "1 0 0 rg" 'v1(1) & " rg"
                    v2(0) = pvRGB(CLng(v1(1))) & " rg"
                Case "b"
            End Select
        Next
        
        Out.Add Array(1, Join(v2, " "))     '-- Add Open Tag
        Out.Add Tmp                         '-- Add Text
        Out.Add Array(2, "")                '-- Add Close Tag

        lastPos = Pos + Len(This.Value)
    Next
    If lastPos <= Len(Text) Then Out.Add Mid$(Text, lastPos)
    Set pvParseTags = Out
End Function
Private Sub pvAddParagraph(Text As String, Out As String, MaxWidth As Single, Justified As Boolean, LineSpacing As Single, OutCy As Single)
Dim Line    As String
Dim LineCx  As Single
Dim Word    As Variant
Dim WordCx  As Single
Dim TestLine As String
    
    For Each Word In Split(Text, " ")
    
        WordCx = TextWidth(IIf(Line = "", Word, " " & Word))
        TestLine = IIf(Line = "", Word, Line & " " & Word)

        If (LineCx + WordCx) <= MaxWidth Then
            Line = TestLine
            LineCx = LineCx + WordCx
        Else
            If Len(Line) > 0 Then
                If Justified Then
                    Out = Out & pvJustify(Line, MaxWidth) & " T*" & vbCrLf
                Else
                    Out = Out & "(" & pvEsc(Line) & ")Tj T*" & vbCrLf
                End If
            End If
            Line = Word
            LineCx = TextWidth(Word)
            OutCy = OutCy + m_Ctx.FontSize + LineSpacing
        End If
    Next
    '-- Add last line
    If Len(Line) > 0 Then
        Out = Out & "(" & pvEsc(Line) & ")Tj" & vbCrLf
        OutCy = OutCy + m_Ctx.FontSize + LineSpacing
    End If
End Sub
Private Sub pvAddParagraphF(Text As String, Out As String, MaxWidth As Single, Justified As Boolean, LineSpacing As Single, OutCy As Single)
Dim Item As Variant
Dim Word As Variant
Dim Buff As String
Dim Line As String
Dim Coll As Collection
Dim bTag    As Boolean
Dim bOpen   As Boolean
Dim bNull   As Boolean
Dim NumSpc  As Long
Dim JBuff   As String

    If Justified Then Set Coll = New Collection
    For Each Item In pvParseTags(Text)
        If Not IsArray(Item) Then
            For Each Word In Split(Item, " ")
                If Len(Line) = 0 Then
                    Line = Word
                ElseIf bTag Then
                    Line = Line & Word
                Else
                    Line = Line & " " & Word
                End If
                If TextWidth(Line) > MaxWidth Then
                    If bOpen Then
                        Buff = Buff & ")Tj"
                        If Justified Then Coll.Add Array("]TJ")
                        bOpen = False
                    End If
                    
                    '-- Break line
                    If Not Justified Then
                        Out = Out & Buff & " T* " & vbCrLf
                    Else
                        Dim t_Width As Single: t_Width = pvTextCx(JBuff)
                        Dim s_Width As Single: s_Width = ((MaxWidth - t_Width) / NumSpc) * (1000 / m_Ctx.FontSize) * -1
                        
                        Dim vm  As Variant
                        Dim Tmp As String
                        
                        For Each vm In Coll
                            If IsArray(vm) Then Tmp = Tmp & vm(0) Else Tmp = Tmp & Str(s_Width) & "(" & vm & ")"
                        Next
                        Out = Out & Tmp & " T*" & vbCrLf
                        Tmp = vbNullString
                        NumSpc = 0
                    End If
                    
                    '-- New Text line at current Word
                    Line = Word
                    JBuff = Word
                    
                    Buff = "(" & pvEsc(CStr(Word))
                    If Justified Then
                        Set Coll = New Collection
                        Coll.Add Array("[")
                        Coll.Add Array("(" & pvEsc(CStr(Word)) & ")")
                    End If
                    bOpen = True
                    OutCy = OutCy + m_Ctx.FontSize + LineSpacing
                Else
                    If Not bOpen Then
                        Buff = Buff & "("
                        If Justified Then Coll.Add Array("[")
                        bOpen = True
                        bTag = True
                    End If
                    Buff = IIf(bTag, Buff & pvEsc(CStr(Word)), Buff & " " & pvEsc(CStr(Word)))
                    If Justified And Len(Word) Then
                        JBuff = JBuff & Word
                        If bTag = True And bNull = False Then
                            Coll.Add Array("(" & pvEsc(CStr(Word)) & ")")
                        ElseIf (Len(Word) > 0) Or bNull Then
                            Coll.Add pvEsc(CStr(Word)): NumSpc = NumSpc + 1
                        End If
                    End If
                End If
                bTag = False
                bNull = (Len(Word) = 0)
            Next
        Else
            Select Case Item(0)
                Case 1: 'Init:  q...(   q...[
                    If bOpen Then '-- Close Previus
                        Buff = Buff & ")Tj"
                        If Justified Then Coll.Add Array("]TJ")
                    End If
                    Buff = Buff & " q " & Item(1)
                    If Justified Then Coll.Add Array(" q " & Item(1))
                    bOpen = False
                    
                Case 2: 'End :  )tj Q   ]TJ Q
                    Buff = Buff & ")Tj Q "
                    If Justified Then Coll.Add Array("]TJ Q ")
                    bOpen = False
            End Select
            bTag = True
        End If
    Next

    '-- Last Line
    If Len(Buff) > 0 Then
        If bOpen Then Buff = Buff & ")Tj" & vbCrLf
        Out = Out & Buff
        OutCy = OutCy + m_Ctx.FontSize + LineSpacing
    End If
End Sub


Private Sub pvInitFontWidths()
On Error GoTo Q:

    ' Index 0 = Chr(32) " ", Index 1 = "!", ..., Index 223 = Chr(255) "ÿ"
    Select Case m_uFonts(m_Ctx.FontKey)
        Case "/Helvetica":
            m_FontWidths = Split( _
            "278|278|355|556|556|889|667|191|333|333|389|584|278|333|278|278|" & _
            "556|556|556|556|556|556|556|556|556|556|278|278|584|584|584|556|" & _
            "1015|667|667|722|722|667|611|778|722|278|500|667|556|833|722|778|" & _
            "667|778|722|667|611|722|667|944|667|667|611|278|278|278|469|556|" & _
            "333|556|556|500|556|556|278|556|556|222|222|500|222|833|556|556|" & _
            "556|556|333|500|278|556|500|722|500|500|500|334|260|334|584|350|" & _
            "556|350|222|556|333|1000|556|556|333|1000|667|333|1000|350|611|350|" & _
            "350|222|222|333|333|350|556|1000|333|1000|500|333|944|350|500|667|" & _
            "278|333|556|556|556|556|260|556|333|737|370|556|584|333|737|333|" & _
            "400|584|333|333|333|556|537|278|333|333|365|556|834|834|834|611|" & _
            "667|667|667|667|667|667|1000|722|667|667|667|667|278|278|278|278|" & _
            "722|722|778|778|778|778|778|584|778|722|722|722|722|667|667|611|" & _
            "556|556|556|556|556|556|889|500|556|556|556|556|278|278|278|278|" & _
            "556|556|556|556|556|556|556|584|611|556|556|556|556|500|556|500", "|")
            m_FontDefWidth = 500
            
        Case "/Helvetica-Bold":
            m_FontWidths = Split( _
            "278|333|474|556|556|889|722|238|333|333|389|584|278|333|278|278|" & _
            "556|556|556|556|556|556|556|556|556|556|333|333|584|584|584|611|" & _
            "975|722|722|722|722|667|611|778|722|278|556|722|611|833|722|778|" & _
            "667|778|722|667|611|722|667|944|667|667|611|333|278|333|584|556|" & _
            "333|556|611|556|611|556|333|611|611|278|278|556|278|889|611|611|" & _
            "611|611|389|556|333|611|556|778|556|556|500|389|280|389|584|350|" & _
            "556|350|278|556|500|1000|556|556|333|1000|667|333|1000|350|611|350|" & _
            "350|278|278|500|500|350|556|1000|333|1000|556|333|944|350|500|667|" & _
            "278|333|556|556|556|556|280|556|333|737|370|556|584|333|737|333|" & _
            "400|584|333|333|333|611|556|278|333|333|365|556|834|834|834|611|" & _
            "722|722|722|722|722|722|1000|722|667|667|667|667|278|278|278|278|" & _
            "722|722|778|778|778|778|778|584|778|722|722|722|722|667|667|611|" & _
            "556|556|556|556|556|556|889|556|556|556|556|556|278|278|278|278|" & _
            "611|611|611|611|611|611|611|584|611|611|611|611|611|556|611|556|", "|")
            m_FontDefWidth = 500
            
        Case "/Helvetica-Oblique":
            m_FontWidths = Split( _
            "278|278|355|556|556|889|667|191|333|333|389|584|278|333|278|278|" & _
            "556|556|556|556|556|556|556|556|556|556|278|278|584|584|584|556|" & _
            "1015|667|667|722|722|667|611|778|722|278|500|667|556|833|722|778|" & _
            "667|778|722|667|611|722|667|944|667|667|611|278|278|278|469|556|" & _
            "333|556|556|500|556|556|278|556|556|222|222|500|222|833|556|556|" & _
            "556|556|333|500|278|556|500|722|500|500|500|334|260|334|584|350|" & _
            "556|350|222|556|333|1000|556|556|333|1000|667|333|1000|350|611|350|" & _
            "350|222|222|333|333|350|556|1000|333|1000|500|333|944|350|500|667|" & _
            "278|333|556|556|556|556|260|556|333|737|370|556|584|333|737|333|" & _
            "400|584|333|333|333|556|537|278|333|333|365|556|834|834|834|611|" & _
            "667|667|667|667|667|667|1000|722|667|667|667|667|278|278|278|278|" & _
            "722|722|778|778|778|778|778|584|778|722|722|722|722|667|667|611|" & _
            "556|556|556|556|556|556|889|500|556|556|556|556|278|278|278|278|" & _
            "556|556|556|556|556|556|556|584|611|556|556|556|556|500|556|500", "|")
            m_FontDefWidth = 500
        Case "/Helvetica-BoldOblique"
            m_FontWidths = Split( _
            "278|333|474|556|556|889|722|238|333|333|389|584|278|333|278|278|" & _
            "556|556|556|556|556|556|556|556|556|556|333|333|584|584|584|611|" & _
            "975|722|722|722|722|667|611|778|722|278|556|722|611|833|722|778|" & _
            "667|778|722|667|611|722|667|944|667|667|611|333|278|333|584|556|" & _
            "333|556|611|556|611|556|333|611|611|278|278|556|278|889|611|611|" & _
            "611|611|389|556|333|611|556|778|556|556|500|389|280|389|584|350|" & _
            "556|350|278|556|500|1000|556|556|333|1000|667|333|1000|350|611|350|" & _
            "350|278|278|500|500|350|556|1000|333|1000|556|333|944|350|500|667|" & _
            "278|333|556|556|556|556|280|556|333|737|370|556|584|333|737|333|" & _
            "400|584|333|333|333|611|556|278|333|333|365|556|834|834|834|611|" & _
            "722|722|722|722|722|722|1000|722|667|667|667|667|278|278|278|278|" & _
            "722|722|778|778|778|778|778|584|778|722|722|722|722|667|667|611|" & _
            "556|556|556|556|556|556|889|556|556|556|556|556|278|278|278|278|" & _
            "611|611|611|611|611|611|611|584|611|611|611|611|611|556|611|556", "|")
            
        Case "/Times-Roman"
            m_FontWidths = Split( _
            "250|333|408|500|500|833|778|180|333|333|500|564|250|333|250|278|" & _
            "500|500|500|500|500|500|500|500|500|500|278|278|564|564|564|444|" & _
            "921|722|667|667|722|611|556|722|722|333|389|722|611|889|722|722|" & _
            "556|722|667|556|611|722|722|944|722|722|611|333|278|333|469|500|" & _
            "333|444|500|444|500|444|333|500|500|278|278|500|278|778|500|500|" & _
            "500|500|333|389|278|500|500|722|500|500|444|480|200|480|541|350|" & _
            "500|350|333|500|444|1000|500|500|333|1000|556|333|889|350|611|350|" & _
            "350|333|333|444|444|350|500|1000|333|980|389|333|722|350|444|722|" & _
            "250|333|500|500|500|500|200|500|333|760|276|500|564|333|760|333|" & _
            "400|564|300|300|333|500|453|250|333|300|310|500|750|750|750|444|" & _
            "722|722|722|722|722|722|889|667|611|611|611|611|333|333|333|333|" & _
            "722|722|722|722|722|722|722|564|722|722|722|722|722|722|556|500|" & _
            "444|444|444|444|444|444|667|444|444|444|444|444|278|278|278|278|" & _
            "500|500|500|500|500|500|500|564|500|500|500|500|500|500|500|500", "|")
            m_FontDefWidth = 333
            
        Case "/Times-Bold"
            m_FontWidths = Split( _
            "250|333|555|500|500|1000|833|278|333|333|500|570|250|333|250|278|" & _
            "500|500|500|500|500|500|500|500|500|500|333|333|570|570|570|500|" & _
            "930|722|667|722|722|667|611|778|778|389|500|778|667|944|722|778|" & _
            "611|778|722|556|667|722|722|1000|722|722|667|333|278|333|581|500|" & _
            "333|500|556|444|556|444|333|500|556|278|333|556|278|833|556|500|" & _
            "556|556|444|389|333|556|500|722|500|500|444|394|220|394|520|350|" & _
            "500|350|333|500|500|1000|500|500|333|1000|556|333|1000|350|667|350|" & _
            "350|333|333|500|500|350|500|1000|333|1000|389|333|722|350|444|722|" & _
            "250|333|500|500|500|500|220|500|333|747|300|500|570|333|747|333|" & _
            "400|570|300|300|333|556|540|250|333|300|330|500|750|750|750|500|" & _
            "722|722|722|722|722|722|1000|722|667|667|667|667|389|389|389|389|" & _
            "722|722|778|778|778|778|778|570|778|722|722|722|722|722|611|556|" & _
            "500|500|500|500|500|500|722|444|444|444|444|444|278|278|278|278|" & _
            "500|556|500|500|500|500|500|570|500|556|556|556|556|500|556|500", "|")
            m_FontDefWidth = 333
            
        Case "/Times-Italic"
        
            m_FontWidths = Split( _
            "250|333|420|500|500|833|778|214|333|333|500|675|250|333|250|278|" & _
            "500|500|500|500|500|500|500|500|500|500|333|333|675|675|675|500|" & _
            "920|611|611|667|722|611|611|722|722|333|444|667|556|833|667|722|" & _
            "611|722|611|500|556|722|611|833|611|556|556|389|278|389|422|500|" & _
            "333|500|500|444|500|444|278|500|500|278|278|444|278|722|500|500|" & _
            "500|500|389|389|278|500|444|667|444|444|389|400|275|400|541|350|" & _
            "500|350|333|500|556|889|500|500|333|1000|500|333|944|350|556|350|" & _
            "350|333|333|556|556|350|500|889|333|980|389|333|667|350|389|556|" & _
            "250|389|500|500|500|500|275|500|333|760|276|500|675|333|760|333|" & _
            "400|675|300|300|333|500|523|250|333|300|310|500|750|750|750|500|" & _
            "611|611|611|611|611|611|889|667|611|611|611|611|333|333|333|333|" & _
            "722|667|722|722|722|722|722|675|722|722|722|722|722|556|611|500|" & _
            "500|500|500|500|500|500|667|444|444|444|444|444|278|278|278|278|" & _
            "500|500|500|500|500|500|500|675|500|500|500|500|500|444|500|444", "|")
            m_FontDefWidth = 389
        Case "/Times-BoldItalic"
            m_FontWidths = Split( _
            "250|389|555|500|500|833|778|278|333|333|500|570|250|333|250|278|" & _
            "500|500|500|500|500|500|500|500|500|500|333|333|570|570|570|500|" & _
            "832|667|667|667|722|667|667|722|778|389|500|667|611|889|722|722|" & _
            "611|722|667|556|611|722|667|889|667|611|611|333|278|333|570|500|" & _
            "333|500|500|444|500|444|333|500|556|278|278|500|278|778|556|500|" & _
            "500|500|389|389|278|556|444|667|500|444|389|348|220|348|570|350|" & _
            "500|350|333|500|500|1000|500|500|333|1000|556|333|944|350|611|350|" & _
            "350|333|333|500|500|350|500|1000|333|1000|389|333|722|350|389|611|" & _
            "250|389|500|500|500|500|220|500|333|747|266|500|606|333|747|333|" & _
            "400|570|300|300|333|576|500|250|333|300|300|500|750|750|750|500|" & _
            "667|667|667|667|667|667|944|667|667|667|667|667|389|389|389|389|" & _
            "722|722|722|722|722|722|722|570|722|722|722|722|722|611|611|500|" & _
            "500|500|500|500|500|500|722|444|444|444|444|444|278|278|278|278|" & _
            "500|556|500|500|500|500|500|570|500|556|556|556|556|444|500|444", "|")
            m_FontDefWidth = 333
            
        Case "/Symbol"
            m_FontWidths = Split( _
            "250|333|713|500|549|833|778|439|333|333|500|549|250|549|250|278|" & _
            "500|500|500|500|500|500|500|500|500|500|278|278|549|549|549|444|" & _
            "549|722|667|722|612|611|763|603|722|333|631|722|686|889|722|722|" & _
            "768|741|556|592|611|690|439|768|645|795|611|333|863|333|658|500|" & _
            "500|631|549|549|494|439|521|411|603|329|603|549|549|576|521|549|" & _
            "549|521|549|603|439|576|713|686|493|686|494|480|200|480|549|0|" & _
            "0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|" & _
            "0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|" & _
            "750|620|247|549|167|713|500|753|753|753|753|1042|987|603|987|603|" & _
            "400|549|411|549|549|713|494|460|549|549|549|549|1000|603|1000|658|" & _
            "823|686|795|987|768|768|823|768|768|713|713|713|713|713|713|713|" & _
            "768|713|790|790|890|823|549|250|713|603|603|1042|987|603|987|603|" & _
            "494|329|790|790|786|713|384|384|384|384|384|384|494|494|494|494|" & _
            "0|329|274|686|686|686|384|384|384|384|384|384|494|494|494|0", "|")
            m_FontDefWidth = 332
            
        Case "/ZapfDingbats"
            m_FontWidths = Split( _
            "278|974|961|974|980|719|789|790|791|690|960|939|549|855|911|933|" & _
            "911|945|974|755|846|762|761|571|677|763|760|759|754|494|552|537|" & _
            "577|692|786|788|788|790|793|794|816|823|789|841|823|833|816|831|" & _
            "923|744|723|749|790|792|695|776|768|792|759|707|708|682|701|826|" & _
            "815|789|789|707|687|696|689|786|787|713|791|785|791|873|761|762|" & _
            "762|759|759|892|892|788|784|438|138|277|415|392|392|668|668|0|" & _
            "390|390|317|317|276|276|509|509|410|410|234|234|334|334|0|0|" & _
            "0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|" & _
            "0|732|544|544|910|667|760|760|776|595|694|626|788|788|788|788|" & _
            "788|788|788|788|788|788|788|788|788|788|788|788|788|788|788|788|" & _
            "788|788|788|788|788|788|788|788|788|788|788|788|788|788|788|788|" & _
            "788|788|788|788|894|838|1016|458|748|924|748|918|927|928|928|834|" & _
            "873|828|924|924|917|930|931|463|883|836|836|867|867|696|696|874|" & _
            "0|874|760|946|771|865|771|888|967|888|831|873|927|970|918|0", "|")
            
        Case "/Courier", "/Courier-Bold", "/Courier-Oblique", "/Courier-BoldOblique"
            ReDim m_FontWidths(0 To 223)
            Dim i As Long
            For i = 0 To UBound(m_FontWidths)
                m_FontWidths(i) = 600
            Next
            m_FontDefWidth = 600
    End Select
Q:
End Sub

Private Function zlib_init(Optional hMod As Long) As Boolean

    If hMod = 0 Then
        hMod = GetModuleHandleW(StrPtr("zlib.dll"))
        If hMod = 0 Then hMod = GetModuleHandleW(StrPtr("zlibwapi.dll"))
        If hMod = 0 Then Exit Function
    End If
    With m_Zlib
        .CompressBound = GetProcAddress(hMod, "compressBound")
        .Compress2 = GetProcAddress(hMod, "compress2")
        If .CompressBound = 0 Or .Compress2 = 0 Then Exit Function
        Const CC_STDCALL = 4&, CC_CDECL = 1&
        If Not IsEmpty(cFunc(.CompressBound, CC_STDCALL, 0&)) Then
            .CConv = CC_STDCALL
        ElseIf Not IsEmpty(cFunc(.CompressBound, CC_CDECL, 0&)) Then
            .CConv = CC_CDECL
        End If
        zlib_init = .CConv
    End With
End Function

Private Function cFunc(lAddr As Long, CC As Integer, ParamArray Params() As Variant) As Variant
Dim vPrms    As Variant
Dim lType()  As Integer
Dim lPtr()   As Long
Dim hResult  As Long
Dim i        As Long

    '/* Copy of Params */
    vPrms = CVar(Params)
    i = UBound(vPrms)
    ReDim lPtr(i)
    ReDim lType(i)
   
    For i = 0 To UBound(vPrms)
        lType(i) = VarType(vPrms(i))
        lPtr(i) = VarPtr(vPrms(i))
    Next
    hResult = DispCallFunc(0&, lAddr, CC, vbLong, i, lType(0), lPtr(0), cFunc)
End Function
