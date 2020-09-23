VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Bevel"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   459
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSplit 
      Height          =   75
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   589
      TabIndex        =   13
      Top             =   3840
      Width           =   8895
   End
   Begin VB.PictureBox picTop 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      HasDC           =   0   'False
      Height          =   1215
      Left            =   120
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   585
      TabIndex        =   3
      Top             =   120
      Width           =   8775
      Begin VB.ComboBox cboPath 
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   0
         Width           =   8055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Go"
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Load"
         Height          =   255
         Left            =   8160
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtLightAngle 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Text            =   "135"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtLightElevation 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Text            =   "30"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtBevelSize 
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         Text            =   "5"
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Debug"
         Height          =   255
         Left            =   5760
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Angle (°)"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Elevation (°)"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Size"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.PictureBox picDebug 
      Height          =   2895
      Left            =   120
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   557
      TabIndex        =   2
      Top             =   3960
      Width           =   8415
   End
   Begin VB.PictureBox picOut 
      Height          =   2535
      Left            =   4200
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin VB.PictureBox picIn 
      Height          =   2535
      Left            =   120
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   0
      Top             =   1320
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private BevelHeight As Long

Private LightElevation As Double
Private LightAngle As Double
Private LightPos As Triplet

Private m_lSplitPerMil As Long  ' Position of splitter in 1000ths
Private m_lSplitDragOff As Long ' Offset from dragging

Private Sub InitLight()
    With LightPos
        .X = Cos(LightElevation) * Cos(LightAngle)
        .Y = -Cos(LightElevation) * Sin(LightAngle)
        .Z = Sin(LightElevation)
    End With
End Sub

Private Sub Command1_Click()
    If Check1.Value = 1 Then
        BevelDebug
    Else
        Bevel
    End If
End Sub


Private Sub Bevel()
    Dim dibSrc As cDIBSection
    Set dibSrc = New cDIBSection
    dibSrc.CreateFromPicture picIn.Picture
    
    Dim arrSrcBytes() As Byte
    Dim tSASrc As SAFEARRAY2D
    
    Dim arrHeight() As Byte, arrWork1() As Byte, arrWork2() As Byte, arrWork3() As Byte
    Dim arrHilite() As Byte
    Dim arrShadow() As Byte
    Dim i As Long, j As Long, k As Long
    Dim w As Long, h As Long
    
    w = dibSrc.Width
    h = dibSrc.Height
    ReDim arrHeight(0 To w - 1, 0 To h - 1)
    ReDim arrWork1(0 To w - 1, 0 To h - 1)
    ReDim arrWork2(0 To w - 1, 0 To h - 1)
    ReDim arrWork3(0 To w - 1, 0 To h - 1)
    ReDim arrHilite(0 To w - 1, 0 To h - 1)
    ReDim arrShadow(0 To w - 1, 0 To h - 1)
    
    ' Setup parameters
    BevelHeight = txtBevelSize.Text
    LightAngle = DegreesToRadians(txtLightAngle.Text)
    LightElevation = DegreesToRadians(txtLightElevation.Text)
    InitLight
    '
    MileStone "1 - Started", True
    
   ' Get all the bits to work on:
   With tSASrc
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = dibSrc.Height
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = dibSrc.BytesPerScanLine
      .pvData = dibSrc.DIBSectionBitsPtr
   End With
   CopyMemory ByVal VarPtrArray(arrSrcBytes()), VarPtr(tSASrc), 4
   
   ' Setup work mask
    For j = 0 To h - 1
        k = 0
        For i = 0 To w - 1
            '
            If arrSrcBytes(k + 1, h - j - 1) = 0 Then
                ' Use black as the transparent colour
                arrHeight(i, j) = 0
                arrWork1(i, j) = 0
            Else
                arrHeight(i, j) = 1
                arrWork1(i, j) = 1
            End If
            '
            k = k + 3
        Next
    Next
    ' Shrink work mask and build height field
    
    MileStone "2 - Building height field"

    For k = 1 To BevelHeight  ' Depth
        For j = 1 To h - 2
            For i = 1 To w - 2
                If arrWork1(i, j) = 0 Then GoTo SET_ZERO
                ' Adjacent pixels:
                If arrWork1(i, j - 1) = 0 Then GoTo SET_ZERO
                If arrWork1(i - 1, j) = 0 Then GoTo SET_ZERO
                If arrWork1(i + 1, j) = 0 Then GoTo SET_ZERO
                If arrWork1(i, j + 1) = 0 Then GoTo SET_ZERO
                arrWork2(i, j) = 1
                arrHeight(i, j) = arrHeight(i, j) + 1
                GoTo RESUME_LOOP
SET_ZERO:
                arrWork2(i, j) = 0
RESUME_LOOP:
            Next
        Next
        'SwapArrayDataPtrs VarPtrArray(arrWork1()), VarPtrArray(arrWork2())
        arrWork1 = arrWork2
    Next
    
    ' At this point, the height field values range from
    ' 0 to (BevelHeight+1)
    ' Normalize them to (0...255)
    MileStone "3 - Normalizing height field"

    NormalizeArray arrHeight, 0, BevelHeight + 1
    
' Blur Height field
    MileStone "4 - Blurring height"

    BlurArray arrHeight

    Dim vx As Triplet, vy As Triplet, vN As Triplet
    Dim vLight As Triplet
    Dim IncidentLight As Double
    
    MileStone "5 - Calculating light"

    ' Calculate incident light
    For j = 1 To h - 2
        For i = 1 To w - 2
            If arrHeight(i, j) = 0 Or arrHeight(i, j) = 255 Then
                ' Do nothing
            Else
                With vN
                    .X = CDbl(arrHeight(i, j)) - 0.25 * (2# * arrHeight(i + 1, j) + arrHeight(i + 1, j - 1) + arrHeight(i + 1, j + 1))
                    .Y = CDbl(arrHeight(i, j)) - 0.25 * (2# * arrHeight(i, j + 1) + arrHeight(i - 1, j + 1) + arrHeight(i + 1, j + 1))
                    '.X = CDbl(arrHeight(i, j)) - arrHeight(i + 1, j)
                    '.Y = CDbl(arrHeight(i, j)) - arrHeight(i, j + 1)
                    .Z = 1
                End With
                IncidentLight = DotTriplet(vN, LightPos) / NormTriplet(vN)
                
                If IncidentLight > 0 Then
                    arrHilite(i, j) = CByte(255 * IncidentLight)
                Else
                    arrShadow(i, j) = CByte(-255 * IncidentLight)
                End If
            End If
        Next
    Next
    MileStone "6 - Blurring lights"

For k = 1 To 3
    BlurArray arrHilite
    BlurArray arrShadow
Next
    
    MileStone "7 - Merging height/light"
    For j = 0 To h - 1
        For i = 0 To w - 1
            arrHilite(i, j) = MulDiv(arrHilite(i, j), 255 - arrHeight(i, j), 255)
        Next
    Next
    For j = 0 To h - 1
        For i = 0 To w - 1
            arrShadow(i, j) = MulDiv(arrShadow(i, j), 255 - arrHeight(i, j), 255)
        Next
    Next
    
    
    MileStone "8 - Rendering"
    ' Render effect
    For j = 0 To h - 1
        k = 0
        For i = 0 To w - 1
            If arrSrcBytes(k + 1, h - j - 1) = 0 Then
' Do nothing
            ElseIf arrHeight(i, j) = 0 Or arrHeight(i, j) = 255 Then
' Do nothing
            Else
                If arrShadow(i, j) > 0 Then
Darken arrShadow(i, j), arrSrcBytes(k + 2, h - 1 - j), arrSrcBytes(k + 1, h - 1 - j), arrSrcBytes(k, h - 1 - j)
                End If
                If arrHilite(i, j) > 0 Then
Lighten arrHilite(i, j), arrSrcBytes(k + 2, h - 1 - j), arrSrcBytes(k + 1, h - 1 - j), arrSrcBytes(k, h - 1 - j)
                End If
            End If
            '
            k = k + 3
        Next
    Next
    MileStone "9 - All done"
    
    dibSrc.PaintPicture picOut.hDC
End Sub

Private Sub BevelDebug()
    Dbg.Clear
    Dbg.SetSpacing 10, 10
    
    Dim dibSrc As cDIBSection
    Set dibSrc = New cDIBSection
    dibSrc.CreateFromPicture picIn.Picture
    
    Dim arrSrcBytes() As Byte
    Dim tSASrc As SAFEARRAY2D
    
    Dim arrHeight() As Byte, arrWork1() As Byte, arrWork2() As Byte
    Dim arrHilite() As Byte
    Dim arrShadow() As Byte
    Dim i As Long, j As Long, k As Long
    Dim w As Long, h As Long
    
    w = dibSrc.Width
    h = dibSrc.Height
    ReDim arrHeight(0 To w - 1, 0 To h - 1)
    ReDim arrWork1(0 To w - 1, 0 To h - 1)
    ReDim arrWork2(0 To w - 1, 0 To h - 1)
    ReDim arrHilite(0 To w - 1, 0 To h - 1)
    ReDim arrShadow(0 To w - 1, 0 To h - 1)
    
    ' Setup parameters
    BevelHeight = txtBevelSize.Text
    LightAngle = DegreesToRadians(txtLightAngle.Text)
    LightElevation = DegreesToRadians(txtLightElevation.Text)
    InitLight
    '
    
   ' Get all the bits to work on:
   With tSASrc
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = dibSrc.Height
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = dibSrc.BytesPerScanLine
      .pvData = dibSrc.DIBSectionBitsPtr
   End With
   CopyMemory ByVal VarPtrArray(arrSrcBytes()), VarPtr(tSASrc), 4
   
   ' Setup work mask
    For j = 0 To h - 1
        k = 0
        For i = 0 To w - 1
            '
            If arrSrcBytes(k + 1, h - j - 1) = 0 Then
                arrHeight(i, j) = 0
                arrWork1(i, j) = 0
            Else
                arrHeight(i, j) = 1
                arrWork1(i, j) = 1
            End If
            '
            k = k + 3
        Next
    Next
    ' Shrink work mask and build height field
    
    Debug.Print "Building height field at " & Timer
    
    For k = 1 To BevelHeight  ' Depth
        For j = 1 To h - 2
            For i = 1 To w - 2
                If arrWork1(i, j) = 0 Then GoTo SET_ZERO
                ' Adjacent pixels:
                If arrWork1(i, j - 1) = 0 Then GoTo SET_ZERO
                If arrWork1(i - 1, j) = 0 Then GoTo SET_ZERO
                If arrWork1(i + 1, j) = 0 Then GoTo SET_ZERO
                If arrWork1(i, j + 1) = 0 Then GoTo SET_ZERO
                arrWork2(i, j) = 1
                arrHeight(i, j) = arrHeight(i, j) + 1
                GoTo RESUME_LOOP
SET_ZERO:
                arrWork2(i, j) = 0
RESUME_LOOP:
            Next
        Next
        'SwapArrayDataPtrs VarPtrArray(arrWork1()), VarPtrArray(arrWork2())
        arrWork1 = arrWork2
    Next
    
    Debug.Print "Normalizing height field at " & Timer

    ' At this point, the height field values range from
    ' 0 to (BevelHeight+1)
    ' Normalize them to (0...255)
    NormalizeArray arrHeight, 0, BevelHeight + 1
    Dbg.NewCell w, h, "Height field"
    For j = 0 To h - 1
        For i = 0 To w - 1
            Dbg.Plot i, j, RGB(arrHeight(i, j), arrHeight(i, j), arrHeight(i, j))
        Next
    Next
    
        Debug.Print "Blurring height field at " & Timer

' Blur Height field
For k = 1 To 1
    Dbg.NewCell w, h, "Blurred height" & k
    BlurArray arrHeight
    For j = 1 To h - 2
        For i = 1 To w - 2
            ' Render height field in grey
            Dbg.Plot i, j, RGB(arrHeight(i, j), arrHeight(i, j), arrHeight(i, j))
        Next
    Next
Next k

    Dim vx As Triplet, vy As Triplet, vN As Triplet
    Dim vLight As Triplet
    Dim IncidentLight As Double

    Debug.Print "Calculating light field at " & Timer

    ' Calculate incident light
    For j = 1 To h - 2
        For i = 1 To w - 2
            If arrHeight(i, j) = 0 Or arrHeight(i, j) = 255 Then
            Else
                With vN
                    .X = CDbl(arrHeight(i, j)) - 0.25 * (2# * arrHeight(i + 1, j) + arrHeight(i + 1, j - 1) + arrHeight(i + 1, j + 1))
                    .Y = CDbl(arrHeight(i, j)) - 0.25 * (2# * arrHeight(i, j + 1) + arrHeight(i - 1, j + 1) + arrHeight(i + 1, j + 1))
                    '.X = CDbl(arrHeight(i, j)) - arrHeight(i + 1, j)
                    '.Y = CDbl(arrHeight(i, j)) - arrHeight(i, j + 1)
                    .Z = 1
                End With
                IncidentLight = DotTriplet(vN, LightPos) / NormTriplet(vN)
                
                If IncidentLight > 0 Then
                    arrHilite(i, j) = CByte(255 * IncidentLight)
                Else
                    arrShadow(i, j) = CByte(-255 * IncidentLight)
                End If
            End If
        Next
    Next
    Dbg.NewCell w, h, "Hilights"
    For j = 0 To h - 1
        For i = 0 To w - 1
            Dbg.Plot i, j, RGB(0, arrHilite(i, j), 255)
        Next
    Next
    Dbg.NewCell w, h, "Shadows"
    For j = 0 To h - 1
        For i = 0 To w - 1
            Dbg.Plot i, j, RGB(255, arrShadow(i, j), 0)
        Next
    Next
        Debug.Print "Blurring light at " & Timer

For k = 1 To 3
    BlurArray arrHilite
    Dbg.NewCell w, h, "Blurred Hilights " & k
    For j = 0 To h - 1
        For i = 0 To w - 1
            Dbg.Plot i, j, RGB(0, arrHilite(i, j), 255)
        Next
    Next
    BlurArray arrShadow
    Dbg.NewCell w, h, "Blurred Shadows " & k
    For j = 0 To h - 1
        For i = 0 To w - 1
            Dbg.Plot i, j, RGB(255, arrShadow(i, j), 0)
        Next
    Next
Next
        Debug.Print "Merging light/height at " & Timer

    Dbg.NewCell w, h, "Hilights + Height"
    For j = 0 To h - 1
        For i = 0 To w - 1
            arrHilite(i, j) = (CLng(arrHilite(i, j)) * (255 - arrHeight(i, j))) \ 255
            Dbg.Plot i, j, RGB(0, arrHilite(i, j), 255)
        Next
    Next
    Dbg.NewCell w, h, "Shadows + Height"
    For j = 0 To h - 1
        For i = 0 To w - 1
            arrShadow(i, j) = (CLng(arrShadow(i, j)) * (255 - arrHeight(i, j))) \ 255
            Dbg.Plot i, j, RGB(255, arrShadow(i, j), 0)
        Next
    Next
    
    Debug.Print "Rendering at " & Timer
    ' Render effect
    For j = 0 To h - 1
        k = 0
        For i = 0 To w - 1
            If arrSrcBytes(k + 1, h - j - 1) = 0 Then
' Do nothing
            ElseIf arrHeight(i, j) = 0 Or arrHeight(i, j) = 255 Then
' Do nothing
            Else
                If arrShadow(i, j) > 0 Then
Darken arrShadow(i, j), arrSrcBytes(k + 2, h - 1 - j), arrSrcBytes(k + 1, h - 1 - j), arrSrcBytes(k, h - 1 - j)
                End If
                If arrHilite(i, j) > 0 Then
Lighten arrHilite(i, j), arrSrcBytes(k + 2, h - 1 - j), arrSrcBytes(k + 1, h - 1 - j), arrSrcBytes(k, h - 1 - j)
                End If
            End If
            '
            k = k + 3
        Next
    Next
        Debug.Print "All done at " & Timer

    
    dibSrc.PaintPicture picOut.hDC
End Sub

Private Sub Lighten(ByVal Amount As Byte, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte)
    'r = CByte(Amount * 255 + (1# - Amount) * r)
    r = MulDiv(r, 255 - Amount, 255) + Amount
    'g = CByte(Amount * 255 + (1# - Amount) * g)
    g = MulDiv(g, 255 - Amount, 255) + Amount
    'b = CByte(Amount * 255 + (1# - Amount) * b)
    b = MulDiv(b, 255 - Amount, 255) + Amount
End Sub
Private Sub Darken(ByVal Amount As Byte, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte)
    'r = CByte((1# - Amount) * r)
    r = MulDiv(r, 255 - Amount, 255)
    'g = CByte((1# - Amount) * g)
    g = MulDiv(g, 255 - Amount, 255)
    'b = CByte((1# - Amount) * b)
    b = MulDiv(b, 255 - Amount, 255)
End Sub

Private Sub Command2_Click()
    picIn.Picture = LoadPicture(cboPath.Text)
End Sub

Private Sub Form_Load()
    Set Dbg.Target = picDebug
    m_lSplitPerMil = 500
    
    cboPath.AddItem App.Path & "\Test2.bmp"
    cboPath.AddItem App.Path & "\Test1.bmp"
    cboPath.Text = App.Path & "\Test2.bmp"
    Command2_Click
End Sub

Private Sub BlurArray(arr() As Byte)
    Dim x1 As Long, y1 As Long
    Dim x2 As Long, y2 As Long
    Dim i As Long, j As Long
    Dim arrTmp() As Byte
    Dim tmp As Long
    x1 = LBound(arr, 1) + 1
    y1 = LBound(arr, 2) + 1
    x2 = UBound(arr, 1) - 1
    y2 = UBound(arr, 2) - 1
    
    ReDim arrTmp(x1 - 1 To x2 + 1, y1 - 1 To y2 + 1)
    
    For j = y1 To y2
        For i = x1 To x2
            ' Centre - coefficient: 1
            tmp = arr(i, j)
            ' Adjacent pixels - coefficient: 3
            tmp = tmp + 4 * arr(i, j - 1)
            tmp = tmp + 4 * arr(i - 1, j)
            tmp = tmp + 4 * arr(i, j + 1)
            tmp = tmp + 4 * arr(i + 1, j)
            ' Diagonal pixels - coefficient: 2
            tmp = tmp + 3 * arr(i - 1, j - 1)
            tmp = tmp + 3 * arr(i + 1, j - 1)
            tmp = tmp + 3 * arr(i - 1, j + 1)
            tmp = tmp + 3 * arr(i + 1, j + 1)
            ' Set light as the weighted average
            arrTmp(i, j) = tmp \ 29
        Next
    Next
    arr = arrTmp
End Sub

Private Sub NormalizeArray(arr() As Byte, ByVal Min As Byte, ByVal Max As Byte)
    Dim x1 As Long, y1 As Long
    Dim x2 As Long, y2 As Long
    Dim i As Long, j As Long
    x1 = LBound(arr, 1): y1 = LBound(arr, 2)
    x2 = UBound(arr, 1): y2 = UBound(arr, 2)
    
    For j = y1 To y2
        For i = x1 To x2
            arr(i, j) = (CLng(arr(i, j) - Min) * 255) \ (Max - Min)
        Next
    Next
End Sub

Private Sub MileStone(sName As String, Optional ByVal Reset As Boolean = False)
    Static LastTime As Single
    If Reset Then
        LastTime = Timer
        Debug.Print sName;
    Else
        Debug.Print 1000 * (Timer - LastTime)
        Debug.Print sName;
        LastTime = Timer
    End If
End Sub

Private Sub Form_Resize()
    Dim AvailH As Long, ViewH As Long
    AvailH = ScaleHeight - picTop.Height
    ViewH = MulDiv(AvailH, m_lSplitPerMil, 1000)
    
    picTop.Move 0, 0, ScaleWidth, picTop.Height
    picIn.Move 0, picTop.Height, ScaleWidth / 2, ViewH
    picOut.Move picIn.Width, picTop.Height, ScaleWidth / 2, ViewH
    picSplit.Move 0, picTop.Height + ViewH, ScaleWidth, picSplit.Height
    picDebug.Move 0, picSplit.Top + picSplit.Height, ScaleWidth, AvailH - ViewH - picSplit.Height
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCapture picSplit.hwnd
    m_lSplitDragOff = Y
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then Exit Sub
    picSplit.Top = picSplit.Top + Y - m_lSplitDragOff
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    m_lSplitPerMil = MulDiv(picSplit.Top - picTop.Height, 1000, ScaleHeight - picTop.Height)
    Form_Resize
End Sub
