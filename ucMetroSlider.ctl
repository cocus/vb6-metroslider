VERSION 5.00
Begin VB.UserControl ucMetroSlider 
   ClientHeight    =   492
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4968
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
End
Attribute VB_Name = "ucMetroSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' =========== Events
Public Event Scroll()
Public Event Change()
Public Event SystemSettingsChanged()


' =========== Internal State
Private Enum eControlState
    [STATE_IDLE] = 0
    [STATE_HOVER] = 1
    [STATE_MOUSE_DOWN] = 2
End Enum

' =========== Private variables
Private u_eState                            As eControlState

Private u_bMouseInControl                   As Boolean

Private u_lMouseScrollLines                 As Long

Private u_hBufferDC                         As Long
Private u_hBitmap                           As Long
Private u_hBitmapOld                        As Long
Private u_hDIB                              As BITMAPINFO

' =========== Local copy of properties
Private u_iSmallChange                      As Integer

Private u_iLargeChange                      As Integer

Private u_dValue                            As Double


' =========== TRACK MOUSE EVENT
Private Enum TRACKMOUSEEVENT_FLAGS
    [TME_HOVER] = &H1&
    [TME_LEAVE] = &H2&
    [TME_QUERY] = &H40000000
    [TME_CANCEL] = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_TYPE
    cbSize      As Long
    dwFlags     As TRACKMOUSEEVENT_FLAGS
    hwndTrack   As Long
    dwHoverTime As Long
End Type

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_TYPE) As Long


' =========== Get MouseWheel Scroll Lines
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const WHEEL_DELTA                   As Long = 120

Private Const SPI_GETWHEELSCROLLLINES       As Long = 104


' =========== Windows Messages
Private Const WM_COMMAND                    As Long = &H111
Private Const WM_MOUSEHOVER                 As Long = &H2A1
Private Const WM_MOUSELEAVE                 As Long = &H2A3
Private Const WM_KILLFOCUS                  As Long = &H8
Private Const WM_SETFOCUS                   As Long = &H7
Private Const WM_PAINT                      As Long = &HF&
Private Const WM_MOUSEACTIVATE              As Long = &H21
Private Const WM_SETFONT                    As Long = &H30
Private Const WM_GETFONT                    As Long = &H31
Private Const WM_KEYDOWN                    As Long = &H100
Private Const WM_KEYUP                      As Long = &H101
Private Const WM_CHAR                       As Long = &H102
Private Const WM_MOUSEMOVE                  As Long = &H200
Private Const WM_LBUTTONUP                  As Long = &H202
Private Const WM_LBUTTONDOWN                As Long = &H201
Private Const WM_RBUTTONDOWN                As Long = &H204
Private Const WM_RBUTTONUP                  As Long = &H205
Private Const WM_MBUTTONDOWN                As Long = &H207
Private Const WM_MBUTTONUP                  As Long = &H208
Private Const WM_MOUSEWHEEL                 As Long = &H20A
Private Const WM_WININICHANGE               As Long = &H1A
Private Const WM_SETTINGCHANGE              As Long = WM_WININICHANGE


' =========== Double Buffer
Private Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type

Private Type BITMAPINFOHEADER
    biSize                  As Long
    biWidth                 As Long
    biHeight                As Long
    biPlanes                As Integer
    biBitCount              As Integer
    biCompression           As Long
    biSizeImage             As Long
    biXPelsPerMeter         As Long
    biYPelsPerMeter         As Long
    biClrUsed               As Long
    biClrImportant          As Long
End Type

Private Type RGBQUAD
    rgbBlue                 As Byte
    rgbGreen                As Byte
    rgbRed                  As Byte
    rgbReserved             As Byte
End Type

Private Type BITMAPINFO
    bmiHeader               As BITMAPINFOHEADER
    bmiColors               As RGBQUAD
End Type

Private Type PAINTSTRUCT
    hdc                                 As Long
    fErase                              As Long
    rcPaint                             As RECT
    fRestore                            As Long
    fIncUpdate                          As Long
    rgbReserved(1 To 32)                As Byte
End Type

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long


' =========== Gdip Load Image + Draw Image
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipDrawImageRect Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipBitmapGetPixel Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByVal mX As Long, ByVal mY As Long, ByRef ARGB As COLORBYTES) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function GdipDrawRectangle Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipFillRectangle Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long

Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal path As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal path As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByVal mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long

Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal Pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long

Private Declare Function GdipFillPath Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal path As Long) As Long


Private Type GdiplusStartupInput
    GDIplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type COLORBYTES
    BlueByte As Byte
    GreenByte As Byte
    RedByte As Byte
    AlphaByte As Byte
End Type

Private Enum SmoothingModes
    SmoothingModeAntiAlias = 4
End Enum

Private Const GWL_WNDPROC       As Long = -4
Private Const GW_OWNER          As Long = 4
Private Const WS_CHILD          As Long = &H40000000
Private Const UnitPixel         As Long = &H2&




' === Subclassing ========================================================
' Subclasing by Paul Caton
Private Enum eMsgWhen                                                      'When to callback
    MSG_BEFORE = 1                                                        'Callback before the original WndProc
    MSG_AFTER = 2                                                         'Callback after the original WndProc
    MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
End Enum

Private Enum eThunkType
    SubclassThunk = 0
    HookThunk = 1
    CallbackThunk = 2
End Enum

Private z_IDEflag           As Long         'Flag indicating we are in IDE
Private z_ScMem             As Long         'Thunk base address
Private z_scFunk            As Collection   'hWnd/thunk-address collection
Private z_hkFunk            As Collection   'hook/thunk-address collection
Private z_cbFunk            As Collection   'callback/thunk-address collection
Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
Private Const IDX_CALLBACKORDINAL As Long = 22 ' Ubound(callback thunkdata)+1, index of the callback

Private Const IDX_WNDPROC   As Long = 9     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
Private Const IDX_UNICODE   As Long = 75    'Must be Ubound(subclass thunkdata)+1; index for unicode support
Private Const ALL_MESSAGES  As Long = -1    'All messages callback
Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SendMessageA Lib "user32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


















Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    UserControl.ForeColor = Value
    Call Redraw
    UserControl.Refresh
    Call PropertyChanged("ForeColor")
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    UserControl.BackColor = Value
    Call Redraw
    UserControl.Refresh
    Call PropertyChanged("BackColor")
End Property

Public Property Get Value() As Double
    Value = u_dValue
End Property

Public Property Let Value(ByVal Value As Double)
    If (Value >= 0) And _
       (Value <= 100) Then
    
        '// Trigger a Scroll event when the value is changed
        If Not (u_dValue = Value) Then
            u_dValue = Value

            Call Redraw
            UserControl.Refresh
            RaiseEvent Scroll
            Call PropertyChanged("Value")
        End If
    End If
End Property










Private Sub UserControl_Initialize()
    u_iSmallChange = 1
    u_iLargeChange = 1
    u_dValue = -1
    
    Call GetScrollLines
    Call CreateBuffer
End Sub

Private Sub UserControl_Terminate()
    Call ssc_Terminate
    Call scb_TerminateCallbacks
    
    Call DisposeBuffer
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", UserControl.BackColor)
        Call .WriteProperty("ForeColor", UserControl.ForeColor)
        Call .WriteProperty("Value", u_dValue)
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UserControl.BackColor = .ReadProperty("BackColor", &H8000000F)
        UserControl.ForeColor = .ReadProperty("ForeColor", &H80000012)
        u_dValue = .ReadProperty("Value", 0)
    End With

    If Ambient.UserMode Then
        Call ManageGDIToken(UserControl.ContainerHwnd)
        
        If ssc_Subclass(UserControl.hwnd) Then
            Call ssc_AddMsg(hwnd, WM_COMMAND, MSG_BEFORE)
            
            Call ssc_AddMsg(hwnd, WM_LBUTTONDOWN, MSG_BEFORE)
            Call ssc_AddMsg(hwnd, WM_RBUTTONDOWN, MSG_BEFORE)
            Call ssc_AddMsg(hwnd, WM_MBUTTONDOWN, MSG_BEFORE)
            
            Call ssc_AddMsg(hwnd, WM_LBUTTONUP, MSG_BEFORE)
            Call ssc_AddMsg(hwnd, WM_RBUTTONUP, MSG_BEFORE)
            Call ssc_AddMsg(hwnd, WM_MBUTTONUP, MSG_BEFORE)

            Call ssc_AddMsg(hwnd, WM_MOUSEMOVE, MSG_BEFORE)
            Call ssc_AddMsg(hwnd, WM_MOUSELEAVE, MSG_BEFORE)
            
            Call ssc_AddMsg(hwnd, WM_MOUSEWHEEL, MSG_BEFORE)
            
            Call ssc_AddMsg(hwnd, WM_SETTINGCHANGE, MSG_BEFORE)
            
            Call ssc_AddMsg(hwnd, WM_PAINT, MSG_BEFORE)
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    Call CreateBuffer
End Sub














Private Sub HandleMouseMove(ByVal iX As Integer, ByVal iY As Integer)
    Dim dNewValue               As Double

    If (iX > UserControl.ScaleWidth) Then
        iX = UserControl.ScaleWidth
    ElseIf (iX < 0) Then
        iX = 0
    End If

    dNewValue = iX / UserControl.ScaleWidth
    
    If Not (dNewValue = u_dValue) Then
        u_dValue = dNewValue

        Call Redraw
        UserControl.Refresh
        RaiseEvent Scroll
    End If
End Sub

Private Sub HandleMouseWheel(ByVal izDelta As Integer)
    Static iAccumDelta          As Integer
    Dim bValueChanged           As Boolean
    
    iAccumDelta = iAccumDelta + izDelta
    
    Do While (iAccumDelta >= u_lMouseScrollLines)
        iAccumDelta = iAccumDelta - u_lMouseScrollLines

        If Not ((u_dValue + (u_iSmallChange / 100)) > 1) Then
            u_dValue = u_dValue + (u_iSmallChange / 100)
            
            bValueChanged = True
        ElseIf Not (u_dValue = 1) Then
            '// Clip to maximum
            u_dValue = 1
            
            bValueChanged = True
        End If
    Loop
    
    Do While (iAccumDelta <= -u_lMouseScrollLines)
        iAccumDelta = iAccumDelta + u_lMouseScrollLines
        
        If Not ((u_dValue - (u_iSmallChange / 100)) < 0) Then
            u_dValue = u_dValue - (u_iSmallChange / 100)
            
            bValueChanged = True
        ElseIf Not (u_dValue = 0) Then
            '// Clip to minimum
            u_dValue = 0
            
            bValueChanged = True
        End If
    Loop
    
    If bValueChanged Then
        Call Redraw
        UserControl.Refresh
        RaiseEvent Scroll
    End If
End Sub










Private Function CreateBuffer() As Boolean
    '// Dispose previous buffer DC and bitmap
    Call DisposeBuffer

    '// Create a new DC
    u_hBufferDC = CreateCompatibleDC(UserControl.hdc)
    
    '// And set its background mode to transparent
    Call SetBkMode(u_hBufferDC, 1)          ' TRANSPARENT

    With u_hDIB.bmiHeader
        .biSize = Len(u_hDIB)
        .biHeight = UserControl.ScaleHeight
        .biWidth = UserControl.ScaleWidth
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0                  ' BI_RGB
    End With
    
    If (SaveDC(u_hBufferDC) = 0) Then
        Exit Function
    End If
    
    u_hBitmap = CreateDIBSection(u_hBufferDC, u_hDIB, 0, 0, 0, 0)
    If (u_hBitmap = 0) Then
        Exit Function
    End If
    
    u_hBitmapOld = SelectObject(u_hBufferDC, u_hBitmap)

    CreateBuffer = Redraw
End Function

Private Sub DisposeBuffer()
    If (u_hBufferDC) Then
        If (u_hBitmapOld) Then
            SelectObject u_hBufferDC, u_hBitmapOld
        End If

        ReleaseDC u_hBufferDC, -1
        DeleteDC u_hBufferDC
    End If

    If (u_hBitmap) Then
        DeleteObject u_hBitmap
    End If
End Sub

Private Function Redraw(Optional ByVal bForce As Boolean = False) As Boolean
    Dim iSliderPosition         As Integer
    Dim hGraphics               As Long

    Dim hSolidBrush             As Long
    Dim rc                      As RECT
    Dim lBackColor              As Long
    Dim lForeColor              As Long
    
    lBackColor = GetTrueColor(BackColor)
    lForeColor = GetTrueColor(ForeColor)

    '// Fill the background of the DC with the selected background color
    hSolidBrush = CreateSolidBrush(lBackColor)
    With rc
        .Left = 0
        .Top = 0
        .Right = UserControl.ScaleWidth
        .Bottom = UserControl.ScaleHeight
    End With
    Call FillRect(u_hBufferDC, rc, hSolidBrush)
    Call DeleteObject(hSolidBrush)
    
    '// Create a GDI+ graphics using the buffer DC
    If GdipCreateFromHDC(u_hBufferDC, hGraphics) = 0 Then
        Call GdipSetInterpolationMode(hGraphics, &H2) 'InterpolationModeHighQuality = &H2

        'Dim iWidth              As Single
        'Dim iHeight             As Single
        Dim hPen                As Long

        'Call GdipGetImageDimension(u_lPictures(u_eState), iWidth, iHeight)
        
        iSliderPosition = (UserControl.ScaleWidth - 9) * u_dValue

        '// Draw the left hand side line
        GdipCreatePen1 ConvertColor(ShiftColor(lForeColor, vbWhite, 200), 100), 3, UnitPixel, hPen
        GdipDrawLine hGraphics, hPen, 0, (UserControl.ScaleHeight / 2), iSliderPosition, (UserControl.ScaleHeight / 2)
        GdipDeletePen hPen

        '// Draw the right hand side line
        GdipCreatePen1 ConvertColor(ShiftColor(lForeColor, &HBBBBBB, 80), 100), 3, UnitPixel, hPen
        GdipDrawLine hGraphics, hPen, iSliderPosition, (UserControl.ScaleHeight / 2), UserControl.ScaleWidth, (UserControl.ScaleHeight / 2)
        GdipDeletePen hPen

        '// Draw the slider
        'GdipDrawImageRect hGraphics, u_lPictures(u_eState), iSliderPosition, 0, iWidth, iHeight

        'GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
        
        Dim mPath               As Long
        Dim hBrush              As Long

        Call GdipCreatePath(&H0, mPath)
        
        GdipAddPathArcI mPath, iSliderPosition - 1, 0, 10, 11, -190, 180
        GdipAddPathArcI mPath, iSliderPosition - 1, (UserControl.ScaleHeight - 12), 10, 11, 0, 180
        
        Call GdipClosePathFigures(mPath)
        
        GdipCreateSolidFill ConvertColor(IIf(u_bMouseInControl, vbWhite, Me.ForeColor), 100), hBrush
        GdipFillPath hGraphics, hBrush, mPath
        
        Call GdipDeleteBrush(hBrush)
        
        Call GdipDeletePath(mPath)
        
        '// Dispose the graphics
        Call GdipDeleteGraphics(hGraphics)
    End If
End Function







Private Sub ChangeState(ByVal eNewState As eControlState)
    If u_eState = eNewState Then
        Exit Sub
    End If
    
    u_eState = eNewState

    Call Redraw
    UserControl.Refresh
End Sub

Private Sub GetScrollLines()
    '// Get how many lines a mouse scroll moves
    Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, u_lMouseScrollLines, 0)
    u_lMouseScrollLines = WHEEL_DELTA / u_lMouseScrollLines
End Sub

Private Function GetTrueColor(ByVal oColor As OLE_COLOR) As OLE_COLOR
    GetTrueColor = oColor
    '// Get the true color if we're using a system constant color
    If (GetTrueColor And &H80000000) Then
        GetTrueColor = GetSysColor(GetTrueColor And &H7FFFFFFF)
    End If
End Function

' funcion para convertir un color long a un BGRA(Blue, Green, Red, Alpha)
Private Function ConvertColor(Color As Long, Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    RtlMoveMemory VarPtr(ConvertColor), VarPtr(BGRA(0)), 4&
End Function

'Funcion para combinar dos colores
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
 
    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
 
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
   
    RtlMoveMemory VarPtr(ShiftColor), VarPtr(clrFore(0)), 4
End Function

Private Sub pvTrackMouseLeave(ByVal lng_hWnd As Long)
    Dim uTME As TRACKMOUSEEVENT_TYPE
    With uTME
        .cbSize = Len(uTME)
        .dwFlags = TME_LEAVE
        .hwndTrack = lng_hWnd
    End With
    Call TrackMouseEvent(uTME)
End Sub







Private Function LoadImageFromFile(ByRef sFileName As String, ByRef hImage As Long) As Boolean
    On Error Resume Next
    Dim FF As Integer, bvStream() As Byte
    FF = FreeFile
    Open sFileName For Binary As #FF
        ReDim bvStream(LOF(FF) - 1)
        Get #FF, , bvStream
    Close #FF
    If Err.Number = 0 Then
        If LoadImageFromStream(bvStream, hImage) Then
            LoadImageFromFile = True
        End If
    End If
End Function

Private Function LoadImageFromStream(ByRef bvData() As Byte, ByRef hImage As Long) As Boolean
    On Local Error GoTo LoadImageFromStream_Error
    Dim IStream     As IUnknown
    
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, hImage) = 0 Then
            LoadImageFromStream = True
        End If
    End If
    Set IStream = Nothing
    
LoadImageFromStream_Error:

End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call RtlMoveMemory(VarPtr(lAddress), lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

'By Lavolpe
Private Function ManageGDIToken(ByVal projectHwnd As Long) As Long
    If projectHwnd = 0& Then Exit Function
    
    Dim hwndGDIsafe     As Long                 'API window to monitor IDE shutdown
    
    Do
        hwndGDIsafe = GetParent(projectHwnd)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    ' ok, got the highest level parent, now find highest level owner
    Do
        hwndGDIsafe = GetWindow(projectHwnd, GW_OWNER)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    
    hwndGDIsafe = FindWindowEx(projectHwnd, 0&, "Static", "GDI+Safe Patch")
    If hwndGDIsafe Then
        ManageGDIToken = hwndGDIsafe    ' we already have a manager running for this VB instance
        Exit Function                   ' can abort
    End If
    
    Dim GDIsi           As GdiplusStartupInput  'GDI+ startup info
    Dim gToken          As Long                 'GDI+ instance token
    
    On Error Resume Next
    GDIsi.GDIplusVersion = 1                    ' attempt to start GDI+
    GdiplusStartup gToken, GDIsi
    If gToken = 0& Then                         ' failed to start
        If Err Then Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    Dim z_ScMem         As Long                 'Thunk base address
    Dim z_Code()        As Long                 'Thunk machine-code initialised here
    Dim nAddr           As Long                 'hwndGDIsafe prev window procedure

    Const WNDPROC_OFF   As Long = &H30          'Offset where window proc starts from z_ScMem
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const MEM_LEN       As Long = &HD4          'Byte length of thunk
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    If z_ScMem <> 0 Then                                     'Ensure the allocation succeeded
        ' we make the api window a child so we can use FindWindowEx to locate it easily
        hwndGDIsafe = CreateWindowExA(0&, "Static", "GDI+Safe Patch", WS_CHILD, 0&, 0&, 0&, 0&, projectHwnd, 0&, App.hInstance, ByVal 0&)
        If hwndGDIsafe <> 0 Then
        
            ReDim z_Code(0 To MEM_LEN \ 4 - 1)
        
            z_Code(12) = &HD231C031: z_Code(13) = &HBBE58960: z_Code(14) = &H12345678: z_Code(15) = &H3FFF631: z_Code(16) = &H74247539: z_Code(17) = &H3075FF5B: z_Code(18) = &HFF2C75FF: z_Code(19) = &H75FF2875
            z_Code(20) = &H2C73FF24: z_Code(21) = &H890853FF: z_Code(22) = &HBFF1C45: z_Code(23) = &H2287D81: z_Code(24) = &H75000000: z_Code(25) = &H443C707: z_Code(26) = &H2&: z_Code(27) = &H2C753339: z_Code(28) = &H2047B81: z_Code(29) = &H75000000
            z_Code(30) = &H2C73FF23: z_Code(31) = &HFFFFFC68: z_Code(32) = &H2475FFFF: z_Code(33) = &H681C53FF: z_Code(34) = &H12345678: z_Code(35) = &H3268&: z_Code(36) = &HFF565600: z_Code(37) = &H43892053: z_Code(38) = &H90909020: z_Code(39) = &H10C261
            z_Code(40) = &H562073FF: z_Code(41) = &HFF2453FF: z_Code(42) = &H53FF1473: z_Code(43) = &H2873FF18: z_Code(44) = &H581053FF: z_Code(45) = &H89285D89: z_Code(46) = &H45C72C75: z_Code(47) = &H800030: z_Code(48) = &H20458B00: z_Code(49) = &H89145D89
            z_Code(50) = &H81612445: z_Code(51) = &H4C4&: z_Code(52) = &HC63FF00

            z_Code(1) = 0                                                   ' shutDown mode; used internally by ASM
            z_Code(2) = zFnAddr("user32", "CallWindowProcA", False)               ' function pointer CallWindowProc
            z_Code(3) = zFnAddr("kernel32", "VirtualFree", False)                  ' function pointer VirtualFree
            z_Code(4) = zFnAddr("kernel32", "FreeLibrary", False)                  ' function pointer FreeLibrary
            z_Code(5) = gToken                                              ' Gdi+ token
            z_Code(10) = LoadLibrary("gdiplus")                             ' library pointer (add reference)
            z_Code(6) = GetProcAddress(z_Code(10), "GdiplusShutdown")       ' function pointer GdiplusShutdown
            z_Code(7) = zFnAddr("user32", "SetWindowLongA", False)                 ' function pointer SetWindowLong
            z_Code(8) = zFnAddr("user32", "SetTimer", False)                       ' function pointer SetTimer
            z_Code(9) = zFnAddr("user32", "KillTimer", False)                      ' function pointer KillTimer
        
            z_Code(14) = z_ScMem                                            ' ASM ebx start point
            z_Code(34) = z_ScMem + WNDPROC_OFF                              ' subclass window procedure location
        
            RtlMoveMemory z_ScMem, VarPtr(z_Code(0)), MEM_LEN               'Copy the thunk code/data to the allocated memory
        
            nAddr = SetWindowLong(hwndGDIsafe, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Subclass our API window
            RtlMoveMemory z_ScMem + 44, VarPtr(nAddr), 4& ' Add prev window procedure to the thunk
            gToken = 0& ' zeroize so final check below does not release it
            
            ManageGDIToken = hwndGDIsafe    ' return handle of our GDI+ manager
        Else
            VirtualFree z_ScMem, 0, MEM_RELEASE     ' failure - release memory
            z_ScMem = 0&
        End If
    Else
        VirtualFree z_ScMem, 0, MEM_RELEASE            ' failure - release memory
        z_ScMem = 0&
    End If
    
    If gToken Then GdiplusShutdown gToken       ' release token if error occurred
    
End Function




'-The following routines are exclusively for the ssc_subclass routines----------------------------
Public Function ssc_Subclass(ByVal lng_hWnd As Long, _
       Optional ByVal lParamUser As Long = 0, _
       Optional ByVal nOrdinal As Long = 1, _
       Optional ByVal oCallback As Object = Nothing, _
       Optional ByVal bIdeSafety As Boolean = True, _
       Optional ByVal bUnicode As Boolean = False) As Boolean 'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
    '* bUnicode - Optional, if True, Unicode API calls will be made to the window vs ANSI calls
    '*************************************************************************************************
    '* cSelfSub - self-subclassing class template
    '* Paul_Caton@hotmail.com
    '* Copyright free, use and abuse as you see fit.
    '*
    '* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
    '* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
    '* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
    '* v1.3 Data array scrapped in favour of property accessors.............................. 20060405
    '* v1.4 Optional IDE protection added
    '*      User-defined callback parameter added
    '*      All user routines that pass in a hWnd get additional validation
    '*      End removed from zError.......................................................... 20060411
    '* v1.5 Added nOrdinal parameter to ssc_Subclass
    '*      Switched machine-code array from Currency to Long................................ 20060412
    '* v1.6 Added an optional callback target object
    '*      Added an IsBadCodePtr on the callback address in the thunk prior to callback..... 20060413
    '*************************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    ' \\LaVolpe - reworked routine a bit, revised the ASM to allow auto-unsubclass on WM_DESTROY
    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    Const CODE_LEN      As Long = 4 * IDX_UNICODE      'Thunk length in bytes
    
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES))  'Bytes to allocate per thunk, data + code + msg tables
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const IDX_EBMODE    As Long = 3                    'Thunk data index of the EbMode function address
    Const IDX_CWP       As Long = 4                    'Thunk data index of the CallWindowProc function address
    Const IDX_SWL       As Long = 5                    'Thunk data index of the SetWindowsLong function address
    Const IDX_FREE      As Long = 6                    'Thunk data index of the VirtualFree function address
    Const IDX_BADPTR    As Long = 7                    'Thunk data index of the IsBadCodePtr function address
    Const IDX_OWNER     As Long = 8                    'Thunk data index of the Owner object's vTable address
    Const IDX_CALLBACK  As Long = 10                   'Thunk data index of the callback method address
    Const IDX_EBX       As Long = 16                   'Thunk code patch index of the thunk data
    Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H38                 'Thunk offset to the WndProc execution address
    Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name
    
    Dim nAddr         As Long
    Dim nID           As Long
    Dim nMyID         As Long

    If IsWindow(lng_hWnd) = 0 Then                      'Ensure the window handle is valid
        Call zError(SUB_NAME, "Invalid window handle")
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                         'Get this process's ID
    Call GetWindowThreadProcessId(lng_hWnd, nID)        'Get the process ID associated with the window handle
    If nID <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        Call zError(SUB_NAME, "Window handle belongs to another process")
        Exit Function
    End If
      
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        Call zError(SUB_NAME, "Callback method not found")
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If z_ScMem <> 0 Then                                  'Ensure the allocation succeeded
        If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
        On Error GoTo CatchDoubleSub                      'Catch double subclassing
        Call z_scFunk.Add(z_ScMem, "h" & lng_hWnd)        'Add the hWnd/thunk-address to the collection
        On Error GoTo 0
        
        ' \\Tai Chi Minh Ralph Eastwood - fixed bug where the MSG_AFTER was not being honored
        ' \\LaVolpe - modified thunks to allow auto-unsubclassing when WM_DESTROY received
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(16) = &H12345678: z_Sc(17) = &HF63103FF: z_Sc(18) = &H750C4339: z_Sc(19) = &H7B8B4A38: z_Sc(20) = &H95E82C: z_Sc(21) = &H7D810000: z_Sc(22) = &H228&: z_Sc(23) = &HC70C7500: z_Sc(24) = &H20443: z_Sc(25) = &H5E90000: z_Sc(26) = &H39000000: z_Sc(27) = &HF751475: z_Sc(28) = &H25E8&: z_Sc(29) = &H8BD23100: z_Sc(30) = &H6CE8307B: z_Sc(31) = &HFF000000: z_Sc(32) = &H10C2610B: z_Sc(33) = &HC53FF00: z_Sc(34) = &H13D&: z_Sc(35) = &H85BE7400: z_Sc(36) = &HE82A74C0: z_Sc(37) = &H2&: z_Sc(38) = &H75FFE5EB: z_Sc(39) = &H2C75FF30: z_Sc(40) = &HFF2875FF: z_Sc(41) = &H73FF2475: z_Sc(42) = &H1053FF24: z_Sc(43) = &H811C4589: z_Sc(44) = &H13B&: z_Sc(45) = &H39727500:
        z_Sc(46) = &H6D740473: z_Sc(47) = &H2473FF58: z_Sc(48) = &HFFFFFC68: z_Sc(49) = &H873FFFF: z_Sc(50) = &H891453FF: z_Sc(51) = &H7589285D: z_Sc(52) = &H3045C72C: z_Sc(53) = &H8000&: z_Sc(54) = &H8920458B: z_Sc(55) = &H4589145D: z_Sc(56) = &HC4816124: z_Sc(57) = &H4&: z_Sc(58) = &H8B1862FF: z_Sc(59) = &H853AE30F: z_Sc(60) = &H810D78C9: z_Sc(61) = &H4C7&: z_Sc(62) = &H28458B00: z_Sc(63) = &H2975AFF2: z_Sc(64) = &H2873FF52: z_Sc(65) = &H5A1C53FF: z_Sc(66) = &H438D1F75: z_Sc(67) = &H144D8D34: z_Sc(68) = &H1C458D50: z_Sc(69) = &HFF3075FF: z_Sc(70) = &H75FF2C75: z_Sc(71) = &H873FF28: z_Sc(72) = &HFF525150: z_Sc(73) = &H53FF2073: z_Sc(74) = &HC328C328
        
        z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
        z_Sc(IDX_INDEX) = lng_hWnd                                               'Store the window handle in the thunk data
        z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
        z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
        
        ' \\LaVolpe - validate unicode request & cache unicode usage
        If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
        z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
        
        ' \\LaVolpe - added extra parameter "bUnicode" to the zFnAddr calls
        z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
        
        Debug.Assert zInIDE
        If bIdeSafety = True And z_IDEflag = 1 Then                   'If the user wants IDE protection
            z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode)    'Store the EbMode function address in the thunk data
        End If
    
        ' \\LaVolpe - use ANSI for non-unicode usage, else use WideChar calls
        If bUnicode Then
            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)          'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)           'Store the SetWindowLong function address in the thunk data
            z_Sc(IDX_UNICODE) = 1
            Call RtlMoveMemory(z_ScMem, VarPtr(z_Sc(0)), CODE_LEN)                  'Copy the thunk code/data to the allocated memory
            nAddr = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
        Else
            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)          'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)           'Store the SetWindowLong function address in the thunk data
            Call RtlMoveMemory(z_ScMem, VarPtr(z_Sc(0)), CODE_LEN)                  'Copy the thunk code/data to the allocated memory
            nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
        End If
        If nAddr = 0 Then                                                           'Ensure the new WndProc was set correctly
            Call zError(SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError)
            GoTo ReleaseMemory
        End If
        'Store the original WndProc address in the thunk data
        Call RtlMoveMemory(z_ScMem + IDX_WNDPROC * 4, VarPtr(nAddr), 4&)        ' z_Sc(IDX_WNDPROC) = nAddr
        ssc_Subclass = True                                                     'Indicate success
    Else
        Call zError(SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError)
    End If

    Exit Function                                                             'Exit ssc_Subclass
    
CatchDoubleSub:
    Call zError(SUB_NAME, "Window handle is already subclassed")
      
ReleaseMemory:
    Call VirtualFree(z_ScMem, 0, MEM_RELEASE)   'ssc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub ssc_Terminate()
    ' can be made public. Releases all subclassing
    ' can be removed and zTerminateThunks can be called directly
    Call zTerminateThunks(SubclassThunk)
End Sub

'UnSubclass the specified window handle
Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public. Releases a specific subclass
    ' can be removed and zUnThunk can be called directly
    Call zUnThunk(lng_hWnd, SubclassThunk)
End Sub

'Add the message value to the window handle's specified callback table
Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    ' Note: can be removed if not needed and zAddMsg can be called directly
    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then                 'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then      'If the message is to be added to the before original WndProc table...
            Call zAddMsg(uMsg, IDX_BTABLE)                  'Add the message to the before table
        End If
        If When And MSG_AFTER Then         'If message is to be added to the after original WndProc table...
            Call zAddMsg(uMsg, IDX_ATABLE)                  'Add the message to the after table
        End If
    End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub ssc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    ' Note: can be removed if not needed and zDelMsg can be called directly
    'Ensure that the thunk hasn't already released its memory
    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then
        'Delete the message from the before table
        If When And MSG_BEFORE Then Call zDelMsg(uMsg, IDX_BTABLE)
        'Delete the message from the after table
        If When And MSG_AFTER Then Call zDelMsg(uMsg, IDX_ATABLE)
    End If
End Sub

'Call the original WndProc
Private Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' Note: can be removed if you do not use this function inside of your window procedure
    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then            'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
End Function

'Get the subclasser lParamUser callback parameter
Private Function zGet_lParamUser(ByVal hWnd_Hook_ID As Long, vType As eThunkType) As Long
    'Note: can be removed if you never need to retrieve/update your user-defined paramter. See ssc_Subclass
    If vType <> CallbackThunk Then
        If IsBadCodePtr(zMap_VFunction(hWnd_Hook_ID, vType)) = 0 Then        'Ensure that the thunk hasn't already released its memory
            zGet_lParamUser = zData(IDX_PARM_USER)                                'Get the lParamUser callback parameter
        End If
    End If
End Function

'Let the subclasser lParamUser callback parameter
Private Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, vType As eThunkType, NewValue As Long)
    'Note: can be removed if you never need to retrieve/update your user-defined paramter. See ssc_Subclass
    If vType <> CallbackThunk Then
        If IsBadCodePtr(zMap_VFunction(hWnd_Hook_ID, vType)) = 0 Then          'Ensure that the thunk hasn't already released its memory
            zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter
        End If
    End If
End Sub

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
    Dim nCount As Long                                                        'Table entry count
    Dim nBase  As Long                                                        'Remember z_ScMem
    Dim i      As Long                                                        'Loop index
    
    nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                   'Map zData() to the specified table
    
    If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
    Else
        nCount = zData(0)                                                       'Get the current table entry count
        If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
            Call zError("zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values")
            GoTo Bail
        End If
    
        For i = 1 To nCount                                                     'Loop through the table entries
            If zData(i) = 0 Then                                                  'If the element is free...
                zData(i) = uMsg                                                     'Use this element
                GoTo Bail                                                           'Bail
            ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
                GoTo Bail                                                           'Bail
            End If
        Next i                                                                  'Next message table entry
    
        nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
        zData(nCount) = uMsg                                                    'Store the message in the appended table entry
    End If
    
    zData(0) = nCount                                                         'Store the new table entry count
Bail:
    z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
    Dim nCount As Long                                                        'Table entry count
    Dim nBase  As Long                                                        'Remember z_ScMem
    Dim i      As Long                                                        'Loop index
    
    nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                   'Map zData() to the specified table
    
    If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0) = 0                                                            'Zero the table entry count
    Else
        nCount = zData(0)                                                       'Get the table entry count
        
        For i = 1 To nCount                                                     'Loop through the table entries
            If zData(i) = uMsg Then                                               'If the message is found...
                zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
                GoTo Bail                                                           'Bail
            End If
        Next i                                                                  'Next message table entry
        
        Call zError("zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table")
    End If
      
Bail:
    z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'-SelfCallback code------------------------------------------------------------------------------------
'-The following routines are exclusively for the scb_SetCallbackAddr routines----------------------------
Private Function scb_SetCallbackAddr(ByVal nParamCount As Long, _
       Optional ByVal nOrdinal As Long = 1, _
       Optional ByVal oCallback As Object = Nothing, _
       Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
    '*************************************************************************************************
    '* nParamCount  - The number of parameters that will callback
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety   - Optional, set to false to disable IDE protection.
    '*************************************************************************************************
    ' Callback procedure must return a Long even if, per MSDN, the callback procedure is a Sub vs Function
    ' The number of parameters are dependent on the individual callback procedures
    
    Const MEM_LEN     As Long = IDX_CALLBACKORDINAL * 4 + 4     'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                         'Commit allocated memory
    Const SUB_NAME      As String = "scb_SetCallbackAddr"       'This routine's name
    Const INDX_OWNER    As Long = 0
    Const INDX_CALLBACK As Long = 1
    Const INDX_EBMODE   As Long = 2
    Const INDX_BADPTR   As Long = 3
    Const INDX_EBX      As Long = 5
    Const INDX_PARAMS   As Long = 12
    Const INDX_PARAMLEN As Long = 17

    Dim z_Cb()    As Long    'Callback thunk array
    Dim nCallback As Long
      
    If z_cbFunk Is Nothing Then
        Set z_cbFunk = New Collection           'If this is the first time through, do the one-time initialization
    Else
        On Error Resume Next                    'Catch already initialized?
        z_ScMem = z_cbFunk.item("h" & nOrdinal) 'Test it
        If Err = 0 Then
            scb_SetCallbackAddr = z_ScMem + 16  'we had this one, just reference it
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    If nParamCount < 0 Then                     ' validate parameters
        Call zError(SUB_NAME, "Invalid Parameter count")
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    nCallback = zAddressOf(oCallback, nOrdinal)         'Get the callback address of the specified ordinal
    If nCallback = 0 Then
        Call zError(SUB_NAME, "Callback address not found.")
        Exit Function
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
        
    If z_ScMem = 0& Then
        Call zError(SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError)  ' oops
        Exit Function
    End If
    Call z_cbFunk.Add(z_ScMem, "h" & nOrdinal)         'Add the callback/thunk-address to the collection
        
    ReDim z_Cb(0 To IDX_CALLBACKORDINAL) As Long       'Allocate for the machine-code array
    
    ' Create machine-code array
    z_Cb(4) = &HBB60E089: z_Cb(6) = &H73FFC589: z_Cb(7) = &HC53FF04: z_Cb(8) = &H7B831F75: z_Cb(9) = &H20750008: z_Cb(10) = &HE883E889: z_Cb(11) = &HB9905004: z_Cb(13) = &H74FF06E3: z_Cb(14) = &HFAE2008D: z_Cb(15) = &H53FF33FF: z_Cb(16) = &HC2906104: z_Cb(18) = &H830853FF: z_Cb(19) = &HD87401F8: z_Cb(20) = &H4589C031: z_Cb(21) = &HEAEBFC
    
    z_Cb(INDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", False)
    z_Cb(INDX_OWNER) = ObjPtr(oCallback)               'Set the Owner
    z_Cb(INDX_CALLBACK) = nCallback                    'Set the callback address
    z_Cb(IDX_CALLBACKORDINAL) = nOrdinal               'Cache ordinal used for zTerminateThunks
      
    Debug.Assert zInIDE
    If bIdeSafety = True And z_IDEflag = 1 Then             'If the user wants IDE protection
        z_Cb(INDX_EBMODE) = zFnAddr("vba6", "EbMode", False)  'EbMode Address
    End If
        
    z_Cb(INDX_PARAMS) = nParamCount                     'Set the parameter count
    z_Cb(INDX_PARAMLEN) = nParamCount * 4               'Set the number of stck bytes to release on thunk return
      
    '\\LaVolpe - redirect address to proper location in virtual memory. Was: z_Cb(INDX_EBX) = VarPtr(z_Cb(INDX_OWNER))
    z_Cb(INDX_EBX) = z_ScMem                            'Set the data address relative to virtual memory pointer
      
    RtlMoveMemory z_ScMem, VarPtr(z_Cb(INDX_OWNER)), MEM_LEN 'Copy thunk code to executable memory
    scb_SetCallbackAddr = z_ScMem + 16                       'Thunk code start address
    
End Function

Private Sub scb_ReleaseCallback(ByVal nOrdinal As Long)
    ' can be made public. Releases a specific callback
    ' can be removed and zUnThunk can be called directly
    Call zUnThunk(nOrdinal, CallbackThunk)
End Sub
Private Sub scb_TerminateCallbacks()
    ' can be made public. Releases all callbacks
    ' can be removed and zTerminateThunks can be called directly
    Call zTerminateThunks(CallbackThunk)
End Sub


'========================================================================
' COMMON USE ROUTINES
'-The following routines are used for each of the three types of thunks
'========================================================================

'Map zData() to the thunk address for the specified window handle
Private Function zMap_VFunction(ByVal vFuncTarget As Long, vType As eThunkType) As Long
    
    ' vFuncTarget is one of the following, depending on vType
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback
    
    Dim thunkCol As Collection
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
    ElseIf vType = HookThunk Then
        Set thunkCol = z_hkFunk
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
    Else
        Call zError("zMap_Vfunction", "Invalid thunk type passed")
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        Call zError("zMap_VFunction", "Thunk hasn't been initialized")
    Else
        On Error GoTo Catch
        z_ScMem = thunkCol("h" & vFuncTarget)                    'Get the thunk address
        zMap_VFunction = z_ScMem
    End If
    Exit Function                                               'Exit returning the thunk address
    
Catch:
    zError "zMap_VFunction", "Thunk type for ID of " & vFuncTarget & " does not exist"
End Function

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
    ' \\LaVolpe -  Note. These two lines can be rem'd out if you so desire. But don't remove the routine
    'App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    'Call MsgBox(sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine)
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
    If asUnicode Then
        zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
    Else
        zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
    End If
    Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
    ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
    Dim bSub  As Byte    'Value we expect to find pointed at by a vTable method entry
    Dim bVal  As Byte
    Dim nAddr As Long    'Address of the vTable
    Dim i     As Long    'Loop index
    Dim J     As Long    'Loop limit
  
    Call RtlMoveMemory(VarPtr(nAddr), ObjPtr(oCallback), 4) 'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then        'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then   'Probe for a Form method
            ' \\LaVolpe - Added propertypage offset
            If Not zProbe(nAddr + &H710, i, bSub) Then  'Probe for a PropertyPage method
                If Not zProbe(nAddr + &H7A4, i, bSub) Then    'Probe for a UserControl method
                    Exit Function                     'Bail...
                End If
            End If
        End If
    End If
  
    i = i + 4                            'Bump to the next entry
    J = i + 1024                         'Set a reasonable limit, scan 256 vTable entries
    Do While i < J
        Call RtlMoveMemory(VarPtr(nAddr), i, 4)  'Get the address stored in this vTable entry
    
        If IsBadCodePtr(nAddr) Then      'Is the entry an invalid code address?
            'Return the specified vTable entry address
            Call RtlMoveMemory(VarPtr(zAddressOf), i - (nOrdinal * 4), 4)
            Exit Do                              'Bad method signature, quit loop
        End If

        Call RtlMoveMemory(VarPtr(bVal), nAddr, 1) 'Get the byte pointed to by the vTable entry
        If bVal <> bSub Then                       'If the byte doesn't match the expected value...
            Call RtlMoveMemory(VarPtr(zAddressOf), i - (nOrdinal * 4), 4) 'Return the specified vTable entry address
            Exit Do                                'Bad method signature, quit loop
        End If
    
        i = i + 4                        'Next vTable entry
    Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
    Dim bVal    As Byte
    Dim nAddr   As Long
    Dim nLimit  As Long
    Dim nEntry  As Long
  
    nAddr = nStart                      'Start address
    nLimit = nAddr + 32                 'Probe eight entries
    Do While nAddr < nLimit             'While we've not reached our probe depth
        Call RtlMoveMemory(VarPtr(nEntry), nAddr, 4)   'Get the vTable entry
    
        If nEntry <> 0 Then                              'If not an implemented interface
            Call RtlMoveMemory(VarPtr(bVal), nEntry, 1)  'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then           'Check for a native or pcode method signature
                nMethod = nAddr                          'Store the vTable entry
                bSub = bVal                              'Store the found method signature
                zProbe = True                            'Indicate success
                Exit Do                                  'Return
            End If
        End If
    
        nAddr = nAddr + 4                                'Next vTable entry
    Loop
End Function

Private Function zInIDE() As Long
    ' This is only run in IDE; it is never run when compiled
    z_IDEflag = 1
    zInIDE = z_IDEflag
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
    ' retrieves stored value from virtual function's memory location
    Call RtlMoveMemory(VarPtr(zData), z_ScMem + (nIndex * 4), 4)
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
    ' sets value in virtual function's memory location
    Call RtlMoveMemory(z_ScMem + (nIndex * 4), VarPtr(nValue), 4)
End Property

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType)
    ' Releases a specific subclass, hook or callback
    ' thunkID depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback

    Const IDX_SHUTDOWN  As Long = 1
    Const MEM_RELEASE As Long = &H8000&                                'Release allocated memory flag
    
    If zMap_VFunction(thunkID, vType) Then
        Select Case vType
            Case SubclassThunk
                If IsBadCodePtr(z_ScMem) = 0 Then          'Ensure that the thunk hasn't already released its memory
                    zData(IDX_SHUTDOWN) = 1                'Set the shutdown indicator
                    Call zDelMsg(ALL_MESSAGES, IDX_BTABLE) 'Delete all before messages
                    Call zDelMsg(ALL_MESSAGES, IDX_ATABLE) 'Delete all after messages
                    '\\LaVolpe - Force thunks to replace original window procedure handle. Without this, app can crash when a window is subclassed multiple times simultaneously
                    If zData(IDX_UNICODE) Then          'Force window procedure handle to be replaced
                        Call SendMessageW(thunkID, 0&, 0&, ByVal 0&)
                    Else
                        Call SendMessageA(thunkID, 0&, 0&, ByVal 0&)
                    End If
                End If
                Call z_scFunk.Remove("h" & thunkID)     'Remove the specified thunk from the collection
            Case HookThunk
                If IsBadCodePtr(z_ScMem) = 0 Then    'Ensure that the thunk hasn't already released its memory
                    zData(IDX_SHUTDOWN) = 1          'Set the shutdown indicator
                    zData(IDX_ATABLE) = 0            'want no more After messages
                    zData(IDX_BTABLE) = 0            'want no more Before messages
                End If
                Call z_hkFunk.Remove("h" & thunkID)           'Remove the specified thunk from the collection
            Case CallbackThunk
                'Ensure that the thunk hasn't already released its memory, and release it
                If IsBadCodePtr(z_ScMem) = 0 Then Call VirtualFree(z_ScMem, 0, MEM_RELEASE)
                Call z_cbFunk.Remove("h" & thunkID)     'Remove the specified thunk from the collection
        End Select
    End If

End Sub

Private Sub zTerminateThunks(ByVal vType As eThunkType)
    ' Removes all thunks of a specific type: subclassing, hooking or callbacks
    Dim i As Long
    Dim thunkCol As Collection
    
    Select Case vType
        Case SubclassThunk
            Set thunkCol = z_scFunk
        Case HookThunk
            Set thunkCol = z_hkFunk
        Case CallbackThunk
            Set thunkCol = z_cbFunk
        Case Else
            Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
        With thunkCol
            For i = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
                z_ScMem = .item(i)                        'Get the thunk address
                If IsBadCodePtr(z_ScMem) = 0 Then         'Ensure that the thunk hasn't already released its memory
                    Select Case vType
                        Case SubclassThunk
                            Call zUnThunk(zData(IDX_INDEX), SubclassThunk)           'Unsubclass
                        Case HookThunk
                            Call zUnThunk(zData(IDX_INDEX), HookThunk)               'Unhook
                        Case CallbackThunk
                            Call zUnThunk(zData(IDX_CALLBACKORDINAL), CallbackThunk) ' release callback
                    End Select
                End If
            Next i                                     'Next member of the collection
        End With
        Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If
End Sub

' === Subclassing callback ===============================================
Private Sub WndProc( _
    ByVal bBefore As Boolean, _
    ByRef bHandled As Boolean, _
    ByRef lReturn As Long, _
    ByVal lng_hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long, _
    ByRef lParamUser As Long)
    
    Dim iX              As Integer
    Dim iY              As Integer
    Dim izDelta         As Integer
    Dim PS              As PAINTSTRUCT
    
    Select Case uMsg
        Case WM_LBUTTONDOWN ', WM_RBUTTONDOWN, WM_MBUTTONDOWN
            '// Mouse Down
            Call ChangeState(STATE_MOUSE_DOWN)

            '// Get mouse X and Y positions from lParam (2 bytes each)
            Call RtlMoveMemory(VarPtr(iX), VarPtr(lParam), 2)
            Call RtlMoveMemory(VarPtr(iY), VarPtr(lParam) + 2, 2)

            '// Handle the movement
            Call HandleMouseMove(iX, iY)

        Case WM_LBUTTONUP ', WM_RBUTTONUP, WM_MBUTTONUP
            If Not u_bMouseInControl Then
                Call ChangeState(STATE_IDLE)
            Else
                Call ChangeState(STATE_HOVER)
            End If

        Case WM_MOUSEWHEEL
            '// Get the delta lines (integer, upper 2 bytes of wParam)
            Call RtlMoveMemory(VarPtr(izDelta), VarPtr(wParam) + 2, 2)

            '// Handle this event
            Call HandleMouseWheel(izDelta)

        Case WM_SETTINGCHANGE
            '// System settings have changed, so get the new value for scroll lines
            Call GetScrollLines
            
            '// And also notify the user about this (maybe to change the colors)
            RaiseEvent SystemSettingsChanged

        Case WM_MOUSEMOVE
            If Not u_bMouseInControl Then
                '// Mouse was not in control, but now it is
                u_bMouseInControl = True

                '// Track mouse leave
                Call pvTrackMouseLeave(lng_hWnd)

                '// Change the state to HOVER
                Call ChangeState(STATE_HOVER)
            Else
                '// If the button is pressed, handle it
                If u_eState = STATE_MOUSE_DOWN Then
                    '// Get mouse X and Y positions from lParam (2 bytes each)
                    Call RtlMoveMemory(VarPtr(iX), VarPtr(lParam), 2)
                    Call RtlMoveMemory(VarPtr(iY), VarPtr(lParam) + 2, 2)

                    '// Handle the movement
                    Call HandleMouseMove(iX, iY)
                End If
            End If

        Case WM_MOUSELEAVE
            '// Mouse is not hovering the control anymore
            u_bMouseInControl = False

            '// Change the state to IDLE
            Call ChangeState(STATE_IDLE)
            
        Case WM_PAINT
            '// Paint event
            If (u_hBufferDC) Then
                '// Handle it only if there is a buffer DC (this should always be the case)

                '// Notify the system that we're paiting now
                Call BeginPaint(hwnd, PS)
                
                '// Draw the buffer DC into the usercontrol DC
                Call BitBlt(UserControl.hdc, _
                            0, 0, _
                            UserControl.Width, UserControl.Height, _
                            u_hBufferDC, _
                            0, 0, _
                            vbSrcCopy)
                
                '// Notify the system that we're done painting
                Call EndPaint(hwnd, PS)
                
                '// And don't allow the UserControl to keep painting it
                bHandled = True
                lReturn = 0
            End If
            
    End Select
End Sub


