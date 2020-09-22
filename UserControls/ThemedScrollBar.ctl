VERSION 5.00
Begin VB.UserControl ThemedScrollBar 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   708
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   864
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ScaleHeight     =   59
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   72
   ToolboxBitmap   =   "ThemedScrollBar.ctx":0000
End
Attribute VB_Name = "ThemedScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ThemedScrollBar Control
'
'Author Ben Vonk
'28-05-2007 First version (Based on Carles P.V. 'ucScrollbar control - version 1.0.4' at http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=63046&lngWId=1)
'08-10-2008 Second version, Add KeyDown, KeyPress and KeyUp events, MouseWheel and ScanMouseWheelInContainer properties for the MouseWheel event

Option Explicit

' Public Events
Public Event Change()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseWheel(ScrollLines As Integer)
Public Event Scroll()

' Private Constants
Private Const ALL_MESSAGES      As Long = -1
Private Const HT_BRBUTTON       As Long = 2
Private Const HT_BRTRACK        As Long = 4
Private Const HT_NOTHING        As Long = 0
Private Const HT_THUMB          As Long = 5
Private Const HT_TLBUTTON       As Long = 1
Private Const HT_TLTRACK        As Long = 3
Private Const MK_LBUTTON        As Long = &H1
Private Const GWL_WNDPROC       As Long = -4
Private Const PATCH_05          As Long = 93
Private Const PATCH_09          As Long = 137
Private Const THUMBSIZE_MIN     As Long = 8
Private Const TIMERID_CHANGE1   As Long = 1
Private Const TIMERID_CHANGE2   As Long = 2
Private Const TIMERID_HOT       As Long = 3
Private Const WM_CANCELMODE     As Long = &H1F
Private Const WM_KEYDOWN        As Long = &H100
Private Const WM_KEYUP          As Long = &H101
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_PAINT          As Long = &HF
Private Const WM_SIZE           As Long = &H5
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_THEMECHANGED   As Long = &H31A
Private Const WM_TIMER          As Long = &H113

' Public Enumeration
Public Enum Orientations
   Horizontal
   Vertical
End Enum

' Private Enumeration
Private Enum MsgWhen
   MSG_BEFORE = 1
   MSG_AFTER = 2
   MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
End Enum

' Private Types
Private Type Bitmap
   bmType                       As Long
   bmWidth                      As Long
   bmHeight                     As Long
   bmWidthBytes                 As Long
   bmPlanes                     As Integer
   bmBitsPixel                  As Integer
   bmBits                       As Long
End Type

Private Type OSVersionInfo
   dwOSVersionInfoSize          As Long
   dwMajorVersion               As Long
   dwMinorVersion               As Long
   dwBuildNumber                As Long
   dwPlatformId                 As Long
   szCSDVersion                 As String * 128
End Type

Private Type PointAPI
   X                            As Long
   Y                            As Long
End Type

Private Type Rect
   Left                         As Long
   Top                          As Long
   Right                        As Long
   Bottom                       As Long
End Type

Private Type PaintStruct
   hDC                          As Long
   fErase                       As Long
   rcPaint                      As Rect
   fRestore                     As Long
   fIncUpdate                   As Long
   rgbReserved(32)              As Byte
End Type

Private Type SubclassDataType
   hWnd                         As Long
   nAddrSclass                  As Long
   nAddrOrig                    As Long
   nMsgCountA                   As Long
   nMsgCountB                   As Long
   aMsgTabelA()                 As Long
   aMsgTabelB()                 As Long
End Type

' Private Variables
Private BRButtonHot             As Boolean
Private BRButtonPressed         As Boolean
Private BRTrackPressed          As Boolean
Private HasNullTrack            As Boolean
Private HasTrack                As Boolean
Private InitProperties          As Boolean
Private IsThemed                As Boolean
Private IsThemedWindows         As Boolean
Private m_ContainerArrowKeys    As Boolean
Private m_MouseWheel            As Boolean
Private m_MouseWheelInContainer As Boolean
Private m_ShowButtons           As Boolean
Private ThumbHot                As Boolean
Private ThumbPressed            As Boolean
Private TLButtonHot             As Boolean
Private TLButtonPressed         As Boolean
Private TLTrackPressed          As Boolean
Private AbsoluteRange           As Long
Private HitTest                 As Long
Private HitTestHot              As Long
Private hPatternBrush           As Long
Private m_LargeChange           As Long
Private m_Max                   As Long
Private m_Min                   As Long
Private m_SmallChange           As Long
Private m_Value                 As Long
Private MouseX                  As Long
Private MouseY                  As Long
Private ScrollLines             As Long
Private ThumbOffset             As Long
Private ThumbPosition           As Long
Private ThumbSize               As Long
Private ValueStartDrag          As Long
Private m_Orientation           As Orientations
Private NullTrackRect           As Rect
Private TLButtonRect            As Rect
Private BRButtonRect            As Rect
Private TLTrackRect             As Rect
Private BRTrackRect             As Rect
Private ThumbRect               As Rect
Private DragRect                As Rect
Private SubclassData()          As SubclassDataType

' Private API's
Private Declare Function CreateBitmap Lib "GDI32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function CreatePatternBrush Lib "GDI32" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "GDI32" (ByVal nIndex As Long) As Long
Private Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function StrLen Lib "Kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function BeginPaint Lib "User32" (ByVal hWnd As Long, lpPaint As PaintStruct) As Long
Private Declare Function CopyRect Lib "User32" (lpDestRect As Rect, lpSourceRect As Rect) As Long
Private Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFrameControl Lib "User32" (ByVal hDC As Long, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function EndPaint Lib "User32" (ByVal hWnd As Long, lpPaint As PaintStruct) As Long
Private Declare Function FillRect Lib "User32" (ByVal hDC As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function GetSysColorBrush Lib "User32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function InflateRect Lib "User32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function InvalidateRect Lib "User32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function KillTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function PtInRect Lib "User32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Private Declare Function SetRect Lib "User32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "User32" (lpRect As Rect) As Long
Private Declare Function SetTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowLongA Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CloseThemeData Lib "UxTheme" (ByVal lngTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "UxTheme" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Rect) As Long
Private Declare Function GetCurrentThemeName Lib "UxTheme" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "UxTheme" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Private Declare Function OpenThemeData Lib "UxTheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub MouseEvents Lib "User32" Alias "mouse_event" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Sub Subclass_WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lhWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Const VK_CONTROL           As Long = &H11
Const VK_MENU              As Long = &H12
Const VK_SHIFT             As Long = &H10

Static intShift            As Integer

Dim lngWindow              As Long
Dim pstPaint               As PaintStruct
Dim ptaMouse               As PointAPI

   Select Case uMsg
      Case WM_KEYDOWN
         Select Case wParam
            Case VK_SHIFT
               intShift = vbShiftMask
               
            Case VK_CONTROL
               intShift = vbCtrlMask
               
            Case VK_MENU
               intShift = vbAltMask
         End Select
         
         If UserControl.Enabled Then If (lhWnd = ContainerHwnd) And m_ContainerArrowKeys Then If (wParam >= vbKeyPageUp) And wParam <= (vbKeyDown) Then Call UserControl_KeyDown(CInt(wParam), intShift)
         
      Case WM_KEYUP
         If wParam >= VK_SHIFT And wParam <= VK_MENU Then intShift = 0
         
      Case WM_LBUTTONDBLCLK
         Call MouseEvents(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
         
      Case WM_LBUTTONDOWN
         Call WhenMouseDown(wParam, lParam)
         
      Case WM_LBUTTONUP, WM_CANCELMODE
         Call WhenMouseUp
         
      Case WM_MOUSEMOVE
         Call WhenMouseMove(wParam, lParam)
         
      Case WM_MOUSEWHEEL
         If UserControl.Enabled Then
            GetCursorPos ptaMouse
            lngWindow = WindowFromPoint(ptaMouse.X, ptaMouse.Y)
            
            If (lngWindow = lhWnd) Or (lngWindow = hWnd) Then
               If Not m_MouseWheel Then
                  Exit Sub
                  
               ElseIf lngWindow <> hWnd Then
                  If (lhWnd = ContainerHwnd) And Not m_MouseWheelInContainer Then Exit Sub
               End If
               
               RaiseEvent MouseWheel(ScrollLines + ((ScrollLines * 2) * (wParam > 0)))
            End If
         End If
         
      Case WM_PAINT
         BeginPaint lhWnd, pstPaint
         
         Call DoPaint(pstPaint.hDC)
         
         EndPaint lhWnd, pstPaint
         bHandled = True
         lReturn = 0
         
      Case WM_SIZE
         Call WhenSize
         
         bHandled = True
         lReturn = 0
         
      Case WM_SYSCOLORCHANGE
         Call SysColorChanged
         
      Case WM_THEMECHANGED
         IsThemed = CheckIsThemed
         InvalidateRect UserControl.hWnd, ByVal 0, 0
         
      Case WM_TIMER
         Call TimerDo(wParam)
   End Select

End Sub

Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long

   Subclass_AddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
   Debug.Assert Subclass_AddrFunc

End Function

Private Function Subclass_Index(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean) As Long

   For Subclass_Index = UBound(SubclassData) To 0 Step -1
      If SubclassData(Subclass_Index).hWnd = lhWnd Then
         If Not bAdd Then Exit Function
         
      ElseIf SubclassData(Subclass_Index).hWnd = 0 Then
         If bAdd Then Exit Function
      End If
   Next 'Subclass_Index
   
   If Not bAdd Then Debug.Assert False

End Function

Private Function Subclass_InIDE() As Boolean

   Debug.Assert Subclass_SetTrue(Subclass_InIDE)

End Function

Private Function Subclass_Initialize(ByVal lhWnd As Long) As Long

Const CODE_LEN                  As Long = 200
Const GMEM_FIXED                As Long = 0
Const PATCH_01                  As Long = 18
Const PATCH_02                  As Long = 68
Const PATCH_03                  As Long = 78
Const PATCH_06                  As Long = 116
Const PATCH_07                  As Long = 121
Const PATCH_0A                  As Long = 186
Const FUNC_CWP                  As String = "CallWindowProcA"
Const FUNC_EBM                  As String = "EbMode"
Const FUNC_SWL                  As String = "SetWindowLongA"
Const MOD_USER                  As String = "User32"
Const MOD_VBA5                  As String = "vba5"
Const MOD_VBA6                  As String = "vba6"

Static bytBuffer(1 To CODE_LEN) As Byte
Static lngCWP                   As Long
Static lngEbMode                As Long
Static lngSWL                   As Long

Dim lngCount                    As Long
Dim lngIndex                    As Long
Dim strHex                      As String

   If bytBuffer(1) Then
      lngIndex = Subclass_Index(lhWnd, True)
      
      If lngIndex = -1 Then
         lngIndex = UBound(SubclassData) + 1
         
         ReDim Preserve SubclassData(lngIndex) As SubclassDataType
      End If
      
      Subclass_Initialize = lngIndex
      
   Else
      strHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
      
      For lngCount = 1 To CODE_LEN
         bytBuffer(lngCount) = Val("&H" & Left(strHex, 2))
         strHex = Mid(strHex, 3)
      Next 'lngCount
      
      If Subclass_InIDE Then
         bytBuffer(16) = &H90
         bytBuffer(17) = &H90
         lngEbMode = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)
         
         If lngEbMode = 0 Then lngEbMode = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)
      End If
      
      lngCWP = Subclass_AddrFunc(MOD_USER, FUNC_CWP)
      lngSWL = Subclass_AddrFunc(MOD_USER, FUNC_SWL)
      
      ReDim SubclassData(0) As SubclassDataType
   End If
   
   With SubclassData(lngIndex)
      .hWnd = lhWnd
      .nAddrSclass = GlobalAlloc(GMEM_FIXED, CODE_LEN)
      .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSclass)
      
      Call CopyMemory(ByVal .nAddrSclass, bytBuffer(1), CODE_LEN)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_01, lngEbMode)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_02, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_03, lngSWL)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_06, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_07, lngCWP)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_0A, ObjPtr(Me))
   End With

End Function

Private Function Subclass_SetTrue(ByRef bValue As Boolean) As Boolean

   Subclass_SetTrue = True
   bValue = True

End Function

Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As MsgWhen = MSG_AFTER)

   With SubclassData(Subclass_Index(lhWnd))
      If When And MSG_BEFORE Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass)
      If When And MSG_AFTER Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass)
   End With

End Sub

Private Sub Subclass_DoAddMsg(ByVal uMsg As Long, ByRef aMsgTabel() As Long, ByRef nMsgCount As Long, ByVal When As MsgWhen, ByVal nAddr As Long)

Const PATCH_04 As Long = 88
Const PATCH_08 As Long = 132

Dim lngEntry   As Long

   ReDim lngOffset(1) As Long
   
   If uMsg = ALL_MESSAGES Then
      nMsgCount = ALL_MESSAGES
      
   Else
      For lngEntry = 1 To nMsgCount - 1
         If aMsgTabel(lngEntry) = 0 Then
            aMsgTabel(lngEntry) = uMsg
            
            GoTo ExitSub
            
         ElseIf aMsgTabel(lngEntry) = uMsg Then
            GoTo ExitSub
         End If
      Next 'lngEntry
      
      nMsgCount = nMsgCount + 1
      
      ReDim Preserve aMsgTabel(1 To nMsgCount) As Long
      
      aMsgTabel(nMsgCount) = uMsg
   End If
   
   If When = MSG_BEFORE Then
      lngOffset(0) = PATCH_04
      lngOffset(1) = PATCH_05
      
   Else
      lngOffset(0) = PATCH_08
      lngOffset(1) = PATCH_09
   End If
   
   If uMsg <> ALL_MESSAGES Then Call Subclass_PatchVal(nAddr, lngOffset(0), VarPtr(aMsgTabel(1)))
   
   Call Subclass_PatchVal(nAddr, lngOffset(1), nMsgCount)
   
ExitSub:
   Erase lngOffset

End Sub

Private Sub Subclass_PatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)

   Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)

End Sub

Private Sub Subclass_PatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)

   Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)

End Sub

Private Sub Subclass_Stop(ByVal lhWnd As Long)

   With SubclassData(Subclass_Index(lhWnd))
      SetWindowLongA .hWnd, GWL_WNDPROC, .nAddrOrig
      
      Call Subclass_PatchVal(.nAddrSclass, PATCH_05, 0)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_09, 0)
      
      GlobalFree .nAddrSclass
      .hWnd = 0
      .nMsgCountA = 0
      .nMsgCountB = 0
      Erase .aMsgTabelA, .aMsgTabelB
   End With

End Sub

Private Sub Subclass_Terminate()

Dim lngCount As Long

   For lngCount = UBound(SubclassData) To 0 Step -1
      If SubclassData(lngCount).hWnd Then Call Subclass_Stop(SubclassData(lngCount).hWnd)
   Next 'lngCount

End Sub

Public Property Get ContainerArrowKeys() As Boolean
Attribute ContainerArrowKeys.VB_Description = "Returns/sets a value that determines whether arrow keys of the control's container will being scannend or not."

   ContainerArrowKeys = m_ContainerArrowKeys

End Property

Public Property Let ContainerArrowKeys(ByVal NewContainerArrowKeys As Boolean)

   m_ContainerArrowKeys = NewContainerArrowKeys
   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

   Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal NewEnable As Boolean)

   UserControl.Enabled = NewEnable
   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Property

Public Property Get hWnd() As Long

   hWnd = UserControl.hWnd

End Property

Public Property Get LargeChange() As Long
Attribute LargeChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks the scroll bar area."

   LargeChange = m_LargeChange

End Property

Public Property Let LargeChange(ByVal NewLargeChange As Long)

   If NewLargeChange < 1 Then NewLargeChange = 1
   
   m_LargeChange = NewLargeChange
   ThumbSize = GetThumbSize
   ThumbPosition = GetThumbPosition
   
   Call SizeTrack
   
   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets a scroll bar position's maximum Value property setting."

   Max = m_Max

End Property

Public Property Let Max(ByVal NewMax As Long)

   If NewMax < m_Min Then NewMax = m_Min
   
   m_Max = NewMax
   AbsoluteRange = m_Max - m_Min
   
   If m_Value > m_Max Then m_Value = m_Max
   
   ThumbSize = GetThumbSize
   ThumbPosition = GetThumbPosition
   
   Call SizeTrack
   
   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Property

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets a scroll bar position's minimum Value property setting."

   Min = m_Min

End Property

Public Property Let Min(ByVal NewMin As Long)

   If NewMin > m_Max Then NewMin = m_Max
   
   m_Min = NewMin
   AbsoluteRange = m_Max - m_Min
   
   If m_Value < m_Min Then m_Value = m_Min
   
   ThumbSize = GetThumbSize
   ThumbPosition = GetThumbPosition
   
   Call SizeTrack
   
   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Property

Public Property Get MouseWheel() As Boolean
Attribute MouseWheel.VB_Description = "Returns/sets a value that determines whether usage of the mouse wheel will being scannend or not."

   MouseWheel = m_MouseWheel

End Property

Public Property Let MouseWheel(ByVal NewScanMouseWheel As Boolean)

   m_MouseWheel = NewScanMouseWheel
   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Property

Public Property Get MouseWheelInContainer() As Boolean
Attribute MouseWheelInContainer.VB_Description = "Returns/sets a value that determines whether usage of the mouse wheel in the control's container area will being scannend or not."

   MouseWheelInContainer = m_MouseWheelInContainer

End Property

Public Property Let MouseWheelInContainer(ByVal NewScanMouseWheelInParent As Boolean)

   m_MouseWheelInContainer = NewScanMouseWheelInParent
   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Property

Public Property Get Orientation() As Orientations
Attribute Orientation.VB_Description = "Returns/sets the Vertical or Horizontal orientation of the scroll bar control."

   Orientation = m_Orientation

End Property

Public Property Let Orientation(ByVal NewOrientation As Orientations)

   If NewOrientation < Horizontal Then
      NewOrientation = Horizontal
      
   ElseIf NewOrientation > Vertical Then
      NewOrientation = Vertical
   End If
   
   m_Orientation = NewOrientation
   
   Call WhenSize

End Property

Public Property Get ShowButtons() As Boolean
Attribute ShowButtons.VB_Description = "Returns/sets a value that determines whether the scroll bar buttons are visible or hidden."

   ShowButtons = m_ShowButtons

End Property

Public Property Let ShowButtons(ByVal NewShowButtons As Boolean)

   m_ShowButtons = NewShowButtons
   
   Call WhenSize

End Property

Public Property Get SmallChange() As Long
Attribute SmallChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks a scroll arrow."

   SmallChange = m_SmallChange

End Property

Public Property Let SmallChange(ByVal NewSmallChange As Long)

   If NewSmallChange < 1 Then NewSmallChange = 1
   
   m_SmallChange = NewSmallChange
   ThumbSize = GetThumbSize
   ThumbPosition = GetThumbPosition
   
   Call SizeTrack
   
   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the value of an object."
Attribute Value.VB_UserMemId = 0

   Value = m_Value

End Property

Public Property Let Value(ByVal NewValue As Long)

Dim lngPrevValue As Long

   If NewValue < m_Min Then
      NewValue = m_Min
      
   ElseIf NewValue > m_Max Then
      NewValue = m_Max
   End If
   
   lngPrevValue = m_Value
   m_Value = NewValue
   ThumbPosition = GetThumbPosition
   
   Call SizeTrack
   
   InvalidateRect UserControl.hWnd, ByVal 0, 0
   
   If m_Value <> lngPrevValue Then RaiseEvent Change

End Property

Public Sub Refresh()

   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Sub

Private Function CheckIsThemed() As Boolean

Const VER_PLATFORM_WIN32_NT As Long = 2

Dim lngLibrary              As Long
Dim osvInfo                 As OSVersionInfo
Dim strTheme                As String
Dim strName                 As String

   IsThemedWindows = False
   
   With osvInfo
      .dwOSVersionInfoSize = Len(osvInfo)
      GetVersionEx osvInfo
      
      If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
         If ((.dwMajorVersion > 4) And .dwMinorVersion) Or (.dwMajorVersion > 5) Then
            IsThemedWindows = True
            lngLibrary = LoadLibrary("UXTheme")
            
            If lngLibrary Then
               strTheme = String(255, vbNullChar)
               GetCurrentThemeName StrPtr(strTheme), Len(strTheme), 0, 0, 0, 0
               strTheme = StripNull(strTheme)
               
               If Len(strTheme) Then
                  strName = String(255, vbNullChar)
                  GetThemeDocumentationProperty StrPtr(strTheme), StrPtr("ThemeName"), StrPtr(strName), Len(strName)
                  CheckIsThemed = (StripNull(strName) <> "")
               End If
               
               FreeLibrary lngLibrary
            End If
         End If
      End If
   End With

End Function

Private Function DrawThemePart(ByVal lhDC As Long, ByVal lPart As Long, ByVal lState As Long, lpRect As Rect) As Boolean

Dim lngTheme As Long

   On Local Error GoTo ExitFunction
   lngTheme = OpenThemeData(UserControl.hWnd, StrPtr("ScrollBar"))
   
   If lngTheme Then
      DrawThemePart = (DrawThemeBackground(lngTheme, lhDC, lPart, lState, lpRect, lpRect) = 0)
      CloseThemeData lngTheme
   End If
   
ExitFunction:
   On Local Error GoTo 0

End Function

Private Function GetScrollButtonSize(ByVal Orientation As Orientations) As Long

Const SM_CXHSCROLL As Long = 21
Const SM_CYVSCROLL As Long = 20

Dim lngIndex       As Long

   If Orientation = Horizontal Then
      lngIndex = SM_CXHSCROLL
      
   Else
      lngIndex = SM_CYVSCROLL
   End If
   
   GetScrollButtonSize = GetSystemMetrics(lngIndex) * -CLng(m_ShowButtons)

End Function

Private Function GetScrollPosition() As Long

   On Local Error Resume Next
   
   If m_Orientation = Horizontal Then
      GetScrollPosition = m_Min + (ThumbPosition - TLButtonRect.Right) / (BRButtonRect.Left - TLButtonRect.Right - ThumbSize) * AbsoluteRange
      
   Else
      GetScrollPosition = m_Min + (ThumbPosition - TLButtonRect.Bottom) / (BRButtonRect.Top - TLButtonRect.Bottom - ThumbSize) * AbsoluteRange
   End If
   
   On Local Error GoTo 0

End Function

Private Function GetThumbPosition() As Long

   On Local Error Resume Next
   
   If m_Orientation = Horizontal Then
      GetThumbPosition = TLButtonRect.Right + (BRButtonRect.Left - TLButtonRect.Right - ThumbSize) / AbsoluteRange * (m_Value - m_Min)
      
      If GetThumbPosition = 0 Then GetThumbPosition = GetScrollButtonSize(Horizontal)
      
   Else
      GetThumbPosition = TLButtonRect.Bottom + (BRButtonRect.Top - TLButtonRect.Bottom - ThumbSize) / AbsoluteRange * (m_Value - m_Min)
      
      If GetThumbPosition = 0 Then GetThumbPosition = GetScrollButtonSize(Vertical)
   End If
   
   On Local Error GoTo 0

End Function

Private Function GetThumbSize() As Long

   On Local Error Resume Next
   
   If m_Orientation = Horizontal Then
      GetThumbSize = (BRButtonRect.Left - TLButtonRect.Right) / (AbsoluteRange / m_LargeChange + 1)
      
   Else
      GetThumbSize = (BRButtonRect.Top - TLButtonRect.Bottom) / (AbsoluteRange / m_LargeChange + 1)
   End If
   
   If GetThumbSize < THUMBSIZE_MIN Then GetThumbSize = THUMBSIZE_MIN
   
   On Local Error GoTo 0

End Function

Private Function ScrollPosDec(ByVal lSteps As Long, Optional ByVal bForceRepaint As Boolean) As Boolean

Dim blnChange    As Boolean
Dim lngPrevValue As Long

   lngPrevValue = m_Value
   m_Value = m_Value - lSteps
   
   If m_Value < m_Min Then m_Value = m_Min
   
   If m_Value <> lngPrevValue Then
      ThumbPosition = GetThumbPosition
      
      Call SizeTrack
      
      blnChange = True
   End If
   
   If blnChange Or bForceRepaint Then
      InvalidateRect UserControl.hWnd, ByVal 0, 0
      
      If blnChange Then RaiseEvent Change
   End If
   
   ScrollPosDec = blnChange

End Function

Private Function ScrollPosInc(ByVal lSteps As Long, Optional ByVal bForceRepaint As Boolean) As Boolean

Dim blnChange    As Boolean
Dim lngPrevValue As Long

   lngPrevValue = m_Value
   m_Value = m_Value + lSteps
    
   If m_Value > m_Max Then m_Value = m_Max
   
   If m_Value <> lngPrevValue Then
      ThumbPosition = GetThumbPosition
      
      Call SizeTrack
      
      blnChange = True
   End If
   
   If blnChange Or bForceRepaint Then
      InvalidateRect UserControl.hWnd, ByVal 0, 0
      
      If blnChange Then RaiseEvent Change
   End If
   
   ScrollPosInc = blnChange

End Function

Private Function StripNull(ByVal Text As String) As String

   StripNull = Left(Text, StrLen(StrPtr(Text)))

End Function

Private Function TestHit(ByVal X As Long, ByVal Y As Long) As Long

   Select Case True
      Case PtInRect(TLButtonRect, X, Y)
         TestHit = HT_TLBUTTON
         
      Case PtInRect(BRButtonRect, X, Y)
         TestHit = HT_BRBUTTON
         
      Case PtInRect(TLTrackRect, X, Y)
         TestHit = HT_TLTRACK
         
      Case PtInRect(BRTrackRect, X, Y)
         TestHit = HT_BRTRACK
         
      Case PtInRect(ThumbRect, X, Y)
         TestHit = HT_THUMB
   End Select

End Function

Private Sub DoPaint(ByVal lhDC As Long)

Const ABS_DOWNDISABLED   As Long = 8
Const ABS_DOWNHOT        As Long = 6
Const ABS_DOWNNORMAL     As Long = 5
Const ABS_DOWNPRESSED    As Long = 7
Const ABS_UPDISABLED     As Long = 4
Const ABS_UPHOT          As Long = 2
Const ABS_UPNORMAL       As Long = 1
Const ABS_UPPRESSED      As Long = 3
Const BDR_RAISED         As Long = &H5
Const BF_RECT            As Long = &HF
Const BLACK_BRUSH        As Long = 4
Const COLOR_BTNFACE      As Long = 15
Const DFC_SCROLL         As Long = 3
Const DFCS_FLAT          As Long = &H4000
Const DFCS_INACTIVE      As Long = &H100
Const DFCS_PUSHED        As Long = &H200
Const DFCS_SCROLLDOWN    As Long = &H1
Const DFCS_SCROLLUP      As Long = &H0
Const GRIPPERSIZE_MIN    As Long = 16
Const HSS_DISABLED       As Long = 4
Const HSS_NORMAL         As Long = 1
Const HSS_PUSHED         As Long = 3
Const SBP_ARROWBTN       As Long = 1
Const SBP_GRIPPERVERT    As Long = 9
Const SBP_LOWERTRACKVERT As Long = 6
Const SBP_THUMBBTNVERT   As Long = 3
Const SBP_UPPERTRACKVERT As Long = 7
Const VSS_DISABLED       As Long = 4
Const VSS_HOT            As Long = 2
Const VSS_NORMAL         As Long = 1
Const VSS_PUSHED         As Long = 3

Dim lngHorizontal        As Long

   lngHorizontal = -CLng(m_Orientation = Horizontal)
   
   If IsThemed Then
      If UserControl.Enabled Then
         If TLButtonHot Then
            DrawThemePart lhDC, SBP_ARROWBTN, ABS_UPHOT + (8 * lngHorizontal), TLButtonRect
            
         Else
            If TLButtonPressed Then
               DrawThemePart lhDC, SBP_ARROWBTN, ABS_UPPRESSED + (8 * lngHorizontal), TLButtonRect
               
            Else
               DrawThemePart lhDC, SBP_ARROWBTN, ABS_UPNORMAL + (8 * lngHorizontal), TLButtonRect
            End If
         End If
         
         If BRButtonHot Then
            DrawThemePart lhDC, SBP_ARROWBTN, ABS_DOWNHOT + (8 * lngHorizontal), BRButtonRect
            
         Else
            If BRButtonPressed Then
               DrawThemePart lhDC, SBP_ARROWBTN, ABS_DOWNPRESSED + (8 * lngHorizontal), BRButtonRect
               
            Else
               DrawThemePart lhDC, SBP_ARROWBTN, ABS_DOWNNORMAL + (8 * lngHorizontal), BRButtonRect
            End If
         End If
         
         If HasTrack Then
            If TLTrackPressed Then
               DrawThemePart lhDC, SBP_UPPERTRACKVERT - (2 * lngHorizontal), HSS_PUSHED, TLTrackRect
               
            Else
               DrawThemePart lhDC, SBP_UPPERTRACKVERT - (2 * lngHorizontal), HSS_NORMAL, TLTrackRect
            End If
            
            If BRTrackPressed Then
               DrawThemePart lhDC, SBP_LOWERTRACKVERT - (2 * lngHorizontal), HSS_PUSHED, BRTrackRect
               
            Else
               DrawThemePart lhDC, SBP_LOWERTRACKVERT - (2 * lngHorizontal), HSS_NORMAL, BRTrackRect
            End If
            
            If ThumbHot Then
               DrawThemePart lhDC, SBP_THUMBBTNVERT - (lngHorizontal), VSS_HOT, ThumbRect
               
               If ThumbSize >= GRIPPERSIZE_MIN Then DrawThemePart lhDC, SBP_GRIPPERVERT - (lngHorizontal), VSS_HOT, ThumbRect
               
            Else
               If ThumbPressed Then
                  DrawThemePart lhDC, SBP_THUMBBTNVERT - (lngHorizontal), VSS_PUSHED, ThumbRect
                  
                  If ThumbSize >= GRIPPERSIZE_MIN Then DrawThemePart lhDC, SBP_GRIPPERVERT - (lngHorizontal), VSS_PUSHED, ThumbRect
                  
               Else
                  DrawThemePart lhDC, SBP_THUMBBTNVERT - (lngHorizontal), VSS_NORMAL, ThumbRect
                  
                  If ThumbSize >= GRIPPERSIZE_MIN Then DrawThemePart lhDC, SBP_GRIPPERVERT - (lngHorizontal), VSS_NORMAL, ThumbRect
               End If
            End If
         End If
         
         If HasNullTrack Then DrawThemePart lhDC, SBP_UPPERTRACKVERT - (2 * lngHorizontal), HSS_NORMAL, NullTrackRect
         
      Else
         DrawThemePart lhDC, SBP_ARROWBTN, ABS_UPDISABLED + (8 * lngHorizontal), TLButtonRect
         DrawThemePart lhDC, SBP_ARROWBTN, ABS_DOWNDISABLED + (8 * lngHorizontal), BRButtonRect
         DrawThemePart lhDC, SBP_UPPERTRACKVERT + (2 * lngHorizontal), HSS_DISABLED, TLTrackRect
         
         If HasTrack Then
            DrawThemePart lhDC, SBP_LOWERTRACKVERT + (2 * lngHorizontal), HSS_DISABLED, BRTrackRect
            DrawThemePart lhDC, SBP_THUMBBTNVERT - (lngHorizontal), VSS_DISABLED, ThumbRect
            
            If ThumbSize >= GRIPPERSIZE_MIN Then DrawThemePart lhDC, SBP_GRIPPERVERT - (lngHorizontal), VSS_DISABLED, ThumbRect
         End If
         
         If HasNullTrack Then DrawThemePart lhDC, SBP_UPPERTRACKVERT - (2 * lngHorizontal), HSS_DISABLED, NullTrackRect
      End If
      
   Else
      If UserControl.Enabled Then
         If TLButtonPressed Then
            DrawFrameControl lhDC, TLButtonRect, DFC_SCROLL, DFCS_SCROLLUP + (2 * lngHorizontal) Or DFCS_FLAT Or DFCS_PUSHED
            
         Else
            DrawFrameControl lhDC, TLButtonRect, DFC_SCROLL, DFCS_SCROLLUP + (2 * lngHorizontal)
         End If
         
         If BRButtonPressed Then
            DrawFrameControl lhDC, BRButtonRect, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lngHorizontal) Or DFCS_FLAT Or DFCS_PUSHED
            
         Else
            DrawFrameControl lhDC, BRButtonRect, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lngHorizontal)
         End If
         
         If HasTrack Then
            If TLTrackPressed Then
               FillRect lhDC, TLTrackRect, GetStockObject(BLACK_BRUSH)
               
            Else
               FillRect lhDC, TLTrackRect, hPatternBrush
            End If
            
            If BRTrackPressed Then
               FillRect lhDC, BRTrackRect, GetStockObject(BLACK_BRUSH)
               
            Else
               FillRect lhDC, BRTrackRect, hPatternBrush
            End If
            
            FillRect lhDC, ThumbRect, GetSysColorBrush(COLOR_BTNFACE)
            DrawEdge lhDC, ThumbRect, BDR_RAISED, BF_RECT
         End If
         
         If HasNullTrack Then FillRect lhDC, NullTrackRect, hPatternBrush
         
      Else
         DrawFrameControl lhDC, TLButtonRect, DFC_SCROLL, DFCS_SCROLLUP + (2 * lngHorizontal) Or DFCS_INACTIVE
         DrawFrameControl lhDC, BRButtonRect, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lngHorizontal) Or DFCS_INACTIVE
         
         If HasTrack Then
            FillRect lhDC, TLTrackRect, hPatternBrush
            FillRect lhDC, BRTrackRect, hPatternBrush
            DrawFrameControl lhDC, ThumbRect, 0, 0
         End If
         
         If HasNullTrack Then FillRect lhDC, NullTrackRect, hPatternBrush
      End If
   End If

End Sub

Private Sub MakePatternBrush()

Dim lngBitmap     As Long
Dim intPattern(7) As Integer

   intPattern(0) = &HAA
   intPattern(1) = &H55
   intPattern(2) = &HAA
   intPattern(3) = &H55
   intPattern(4) = &HAA
   intPattern(5) = &H55
   intPattern(6) = &HAA
   intPattern(7) = &H55
   lngBitmap = CreateBitmap(8, 8, 1, 1, intPattern(0))
   hPatternBrush = CreatePatternBrush(lngBitmap)
   DeleteObject lngBitmap

End Sub

Private Sub MakePoints(ByVal IsPoint As Long, X As Long, Y As Long)

   If IsPoint And &H8000& Then
      X = &H8000 Or (IsPoint And &H7FFF&)
      
   Else
      X = IsPoint And &HFFFF&
   End If
   
   If IsPoint And &H80000000 Then
      Y = (IsPoint \ &H10000) - 1
      
   Else
      Y = IsPoint \ &H10000
   End If

End Sub

Private Sub SizeButtons()

Dim lngButtonSize  As Long
Dim lngX           As Long
Dim lngY           As Long
Dim rctClient      As Rect

   GetClientRect hWnd, rctClient
   HasTrack = False
   HasNullTrack = False
   
   If m_Orientation = Horizontal Then
      lngButtonSize = GetScrollButtonSize(Horizontal)
      
      With rctClient
         If 2 * lngButtonSize + THUMBSIZE_MIN > .Right Then
            If 2 * lngButtonSize < .Right Then
               SetRect TLButtonRect, 0, 0, lngButtonSize, .Bottom
               SetRect BRButtonRect, .Right - lngButtonSize, 0, .Right, .Bottom
               HasNullTrack = True
               SetRect NullTrackRect, lngButtonSize, 0, .Right - lngButtonSize, .Bottom
               
            Else
               SetRect TLButtonRect, 0, 0, .Right \ 2, .Bottom
               SetRect BRButtonRect, .Right \ 2 + (.Right Mod 2), 0, .Right, .Bottom
               HasNullTrack = CBool(.Right Mod 2)
               
               If HasNullTrack Then SetRect NullTrackRect, .Right \ 2, 0, .Right \ 2 + 1, .Bottom
            End If
            
         Else
            HasTrack = True
            SetRect TLButtonRect, 0, 0, lngButtonSize, .Bottom
            SetRect BRButtonRect, .Right - lngButtonSize, 0, .Right, .Bottom
         End If
      End With
      
      lngX = 25
      lngY = 250
      
   Else
      lngButtonSize = GetScrollButtonSize(Vertical)
      
      With rctClient
         If 2 * lngButtonSize + THUMBSIZE_MIN > .Bottom Then
            If 2 * lngButtonSize < .Bottom Then
               SetRect TLButtonRect, 0, 0, .Right, lngButtonSize
               SetRect BRButtonRect, 0, .Bottom - lngButtonSize, .Right, .Bottom
               HasNullTrack = True
               SetRect NullTrackRect, 0, lngButtonSize, .Right, .Bottom - lngButtonSize
               
            Else
               SetRect TLButtonRect, 0, 0, .Right, .Bottom \ 2
               SetRect BRButtonRect, 0, .Bottom \ 2 + (.Bottom Mod 2), .Right, .Bottom
               HasNullTrack = CBool(.Bottom Mod 2)
               
               If HasNullTrack Then SetRect NullTrackRect, 0, .Bottom \ 2, .Right, .Bottom \ 2 + 1
            End If
            
         Else
            HasTrack = True
            SetRect TLButtonRect, 0, 0, .Right, lngButtonSize
            SetRect BRButtonRect, 0, .Bottom - lngButtonSize, .Right, .Bottom
         End If
      End With
      
      lngX = 250
      lngY = 25
   End If
   
   CopyRect DragRect, rctClient
   InflateRect DragRect, lngX, lngY
   
   If Not HasTrack Then
      SetRectEmpty TLTrackRect
      SetRectEmpty BRTrackRect
      SetRectEmpty ThumbRect
   End If

End Sub

Private Sub SizeTrack()

   If HasTrack Then
      If m_Orientation = Horizontal Then
         SetRect TLTrackRect, TLButtonRect.Right, 0, ThumbPosition, TLButtonRect.Bottom
         SetRect BRTrackRect, ThumbPosition + ThumbSize, 0, BRButtonRect.Left, BRButtonRect.Bottom
         SetRect ThumbRect, ThumbPosition, 0, ThumbPosition + ThumbSize, BRButtonRect.Bottom
         
      Else
         SetRect TLTrackRect, 0, TLButtonRect.Bottom, TLButtonRect.Right, ThumbPosition
         SetRect BRTrackRect, 0, ThumbPosition + ThumbSize, BRButtonRect.Right, BRButtonRect.Top
         SetRect ThumbRect, 0, ThumbPosition, BRButtonRect.Right, ThumbPosition + ThumbSize
      End If
   End If

End Sub

Private Sub SysColorChanged()

   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Sub

Private Sub TimerDo(ByVal wParam As Long)

Dim ptaMouseXY As PointAPI

   If wParam = TIMERID_CHANGE1 Then
      Call TimerKill(TIMERID_CHANGE1)
      Call TimerSet(TIMERID_CHANGE2, 25)
      
   ElseIf wParam = TIMERID_CHANGE2 Then
      If HitTest = HT_TLBUTTON Then
         If PtInRect(TLButtonRect, MouseX, MouseY) And Not ScrollPosDec(m_SmallChange) Then Call TimerKill(TIMERID_CHANGE2)
         
      ElseIf HitTest = HT_BRBUTTON Then
         If PtInRect(BRButtonRect, MouseX, MouseY) And Not ScrollPosInc(m_SmallChange) Then Call TimerKill(TIMERID_CHANGE2)
         
      ElseIf HitTest = HT_TLTRACK Then
         If m_Orientation = Horizontal Then
            If ThumbPosition > MouseX Then
               TLTrackPressed = True
               ScrollPosDec m_LargeChange
               
            Else
               TLTrackPressed = False
               InvalidateRect UserControl.hWnd, ByVal 0, 0
            End If
            
         Else
            If ThumbPosition > MouseY Then
               TLTrackPressed = True
               ScrollPosDec m_LargeChange
               
            Else
               TLTrackPressed = False
               InvalidateRect UserControl.hWnd, ByVal 0, 0
            End If
         End If
         
      ElseIf HitTest = HT_BRTRACK Then
         If m_Orientation = Horizontal Then
            If ThumbPosition + ThumbSize < MouseX Then
               BRTrackPressed = True
               ScrollPosInc m_LargeChange
               
            Else
               BRTrackPressed = False
               InvalidateRect UserControl.hWnd, ByVal 0, 0
            End If
            
         Else
            If ThumbPosition + ThumbSize < MouseY Then
               BRTrackPressed = True
               ScrollPosInc m_LargeChange
               
            Else
               BRTrackPressed = False
               InvalidateRect UserControl.hWnd, ByVal 0, 0
            End If
         End If
      End If
      
   ElseIf wParam = TIMERID_HOT Then
      GetCursorPos ptaMouseXY
      ScreenToClient hWnd, ptaMouseXY
      
      Select Case True
         Case TLButtonHot
            If PtInRect(TLButtonRect, ptaMouseXY.X, ptaMouseXY.Y) = 0 Then
               TLButtonHot = False
               
               Call TimerKill(TIMERID_HOT)
               
               InvalidateRect UserControl.hWnd, ByVal 0, 0
            End If
            
         Case BRButtonHot
            If PtInRect(BRButtonRect, ptaMouseXY.X, ptaMouseXY.Y) = 0 Then
               BRButtonHot = False
               
               Call TimerKill(TIMERID_HOT)
               
               InvalidateRect UserControl.hWnd, ByVal 0, 0
            End If
            
         Case ThumbHot
            If PtInRect(ThumbRect, ptaMouseXY.X, ptaMouseXY.Y) = 0 Then
               ThumbHot = False
               
               Call TimerKill(TIMERID_HOT)
               
               InvalidateRect UserControl.hWnd, ByVal 0, 0
            End If
      End Select
   End If

End Sub

Private Sub TimerKill(ByVal TimerID As Long)

   KillTimer UserControl.hWnd, TimerID
   HitTestHot = HT_NOTHING

End Sub

Private Sub TimerSet(ByVal TimerID As Long, ByVal DelayTime As Long)

   SetTimer UserControl.hWnd, TimerID, DelayTime, 0

End Sub

Private Sub WhenMouseDown(ByVal wParam As Long, ByVal lParam As Long)

   If wParam And (MK_LBUTTON = MK_LBUTTON) Then
      Call MakePoints(lParam, MouseX, MouseY)
      
      HitTest = TestHit(MouseX, MouseY)
      
      If HitTest = HT_THUMB Then
         If m_Orientation = Horizontal Then
            ThumbOffset = ThumbRect.Left - MouseX
            
         Else
            ThumbOffset = ThumbRect.Top - MouseY
         End If
         
         ThumbPressed = True
         ThumbHot = False
         ValueStartDrag = m_Value
         InvalidateRect UserControl.hWnd, ByVal 0, 0
          
      ElseIf HitTest = HT_TLBUTTON Then
         TLButtonPressed = True
         TLButtonHot = False
         ScrollPosDec m_SmallChange, True
         
         Call TimerKill(TIMERID_CHANGE1)
         Call TimerSet(TIMERID_CHANGE1, 300)
         
      ElseIf HitTest = HT_BRBUTTON Then
         BRButtonPressed = True
         BRButtonHot = False
         ScrollPosInc m_SmallChange, True
         
         Call TimerKill(TIMERID_CHANGE1)
         Call TimerSet(TIMERID_CHANGE1, 300)
         
      ElseIf HitTest = HT_TLTRACK Then
         TLTrackPressed = True
         ScrollPosDec m_LargeChange
         
         Call TimerKill(TIMERID_CHANGE1)
         Call TimerSet(TIMERID_CHANGE1, 300)
         
      ElseIf HitTest = HT_BRTRACK Then
         BRTrackPressed = True
         ScrollPosInc m_LargeChange
         
         Call TimerKill(TIMERID_CHANGE1)
         Call TimerSet(TIMERID_CHANGE1, 300)
      End If
   End If

End Sub

Private Sub WhenMouseMove(ByVal wParam As Long, ByVal lParam As Long)

Const TIMERDT_HOT   As Long = 25

Dim blnHot          As Boolean
Dim blnPressed      As Boolean
Dim lngPrevThumbPos As Long
Dim lngPrevValue    As Long

   Call MakePoints(lParam, MouseX, MouseY)
   
   If wParam And (MK_LBUTTON = MK_LBUTTON) Then
      If HitTest = HT_THUMB Then
         lngPrevValue = m_Value
         lngPrevThumbPos = ThumbPosition
         
         If PtInRect(DragRect, MouseX, MouseY) Then
            If m_Orientation = Horizontal Then
               ThumbPosition = MouseX + ThumbOffset
               
               If ThumbPosition < TLButtonRect.Right Then ThumbPosition = TLButtonRect.Right
               If ThumbPosition + ThumbSize > BRButtonRect.Left Then ThumbPosition = BRButtonRect.Left - ThumbSize
               
            Else
               ThumbPosition = MouseY + ThumbOffset
               
               If ThumbPosition < TLButtonRect.Bottom Then ThumbPosition = TLButtonRect.Bottom
               If ThumbPosition + ThumbSize > BRButtonRect.Top Then ThumbPosition = BRButtonRect.Top - ThumbSize
            End If
            
            m_Value = GetScrollPosition
            
         Else
            m_Value = ValueStartDrag
            ThumbPosition = GetThumbPosition
         End If
         
         If ThumbPosition <> lngPrevThumbPos Then
            Call SizeTrack
            
            InvalidateRect UserControl.hWnd, ByVal 0, 0
            
            If m_Value <> lngPrevValue Then RaiseEvent Scroll
         End If
         
      ElseIf HitTest = HT_TLBUTTON Then
         blnPressed = PtInRect(TLButtonRect, MouseX, MouseY)
         
         If blnPressed Xor TLButtonPressed Then
            TLButtonPressed = blnPressed
            InvalidateRect UserControl.hWnd, ByVal 0, 0
         End If
         
      ElseIf HitTest = HT_BRBUTTON Then
         blnPressed = PtInRect(BRButtonRect, MouseX, MouseY)
         
         If blnPressed Xor BRButtonPressed Then
            BRButtonPressed = blnPressed
            InvalidateRect UserControl.hWnd, ByVal 0, 0
         End If
      End If
      
   Else
      HitTestHot = TestHit(MouseX, MouseY)
      
      If HitTestHot = HT_TLBUTTON Then
         blnHot = PtInRect(TLButtonRect, MouseX, MouseY)
         
         If TLButtonHot Xor blnHot Then
            TLButtonHot = True
            BRButtonHot = False
            ThumbHot = False
            InvalidateRect UserControl.hWnd, ByVal 0, 0
            
            Call TimerKill(TIMERID_HOT)
            Call TimerSet(TIMERID_HOT, TIMERDT_HOT)
         End If
         
      ElseIf HitTestHot = HT_BRBUTTON Then
         blnHot = PtInRect(BRButtonRect, MouseX, MouseY)
         
         If BRButtonHot Xor blnHot Then
            TLButtonHot = False
            BRButtonHot = True
            ThumbHot = False
            InvalidateRect UserControl.hWnd, ByVal 0, 0
            
            Call TimerKill(TIMERID_HOT)
            Call TimerSet(TIMERID_HOT, TIMERDT_HOT)
         End If
         
      ElseIf HitTestHot = HT_THUMB Then
         blnHot = PtInRect(ThumbRect, MouseX, MouseY)
         
         If ThumbHot Xor blnHot Then
            TLButtonHot = False
            BRButtonHot = False
            ThumbHot = True
            InvalidateRect UserControl.hWnd, ByVal 0, 0
            
            Call TimerKill(TIMERID_HOT)
            Call TimerSet(TIMERID_HOT, TIMERDT_HOT)
         End If
      End If
   End If
   
   DoEvents

End Sub

Private Sub WhenMouseUp()

   Call TimerKill(TIMERID_HOT)
   Call TimerKill(TIMERID_CHANGE1)
   Call TimerKill(TIMERID_CHANGE2)
   
   If HitTest = HT_THUMB Then If m_Value <> ValueStartDrag Then RaiseEvent Change
   
   HitTest = HT_NOTHING
   TLButtonPressed = False
   BRButtonPressed = False
   ThumbPressed = False
   TLTrackPressed = False
   BRTrackPressed = False
   ThumbPosition = GetThumbPosition
   
   Call SizeTrack
   
   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Sub

Private Sub WhenSize()

   Call SizeButtons
   
   ThumbSize = GetThumbSize
   ThumbPosition = GetThumbPosition
   
   Call SizeTrack
   
   InvalidateRect UserControl.hWnd, ByVal 0, 0

End Sub

Private Sub UserControl_Initialize()

   IsThemed = CheckIsThemed
   
   Call MakePatternBrush

End Sub

Private Sub UserControl_InitProperties()

   InitProperties = True
   m_LargeChange = 1
   m_MouseWheel = True
   m_ShowButtons = True
   m_SmallChange = 1
   
   Call SizeButtons
   
   ThumbSize = GetThumbSize
   ThumbPosition = GetThumbPosition
   
   Call SizeTrack

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Paint()

   If Not Ambient.UserMode Then Call DoPaint(UserControl.hDC)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Const SPI_GETWHEELSCROLLLINES As Long = 104

   With PropBag
      m_ContainerArrowKeys = .ReadProperty("ContainerArrowKeys", False)
      UserControl.Enabled = .ReadProperty("Enabled", True)
      m_LargeChange = .ReadProperty("LargeChange", 1)
      m_Max = .ReadProperty("Max", 0)
      m_Min = .ReadProperty("Min", 0)
      m_MouseWheel = .ReadProperty("MouseWheel", True)
      m_MouseWheelInContainer = .ReadProperty("MouseWheelInContainer", False)
      m_ShowButtons = .ReadProperty("ShowButtons", True)
      m_SmallChange = .ReadProperty("SmallChange", 1)
      m_Orientation = .ReadProperty("Orientation", Vertical)
      m_Value = .ReadProperty("Value", 0)
   End With
   
   AbsoluteRange = m_Max - m_Min
   SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, ScrollLines, 0
   ScrollLines = ScrollLines + (1 And (ScrollLines = 0))
   
   Call SizeButtons
   
   ThumbSize = GetThumbSize
   ThumbPosition = GetThumbPosition
   
   Call SizeTrack
   
   If Ambient.UserMode Then
      With UserControl
         IsThemed = CheckIsThemed
         
         Call Subclass_Initialize(.hWnd)
         Call Subclass_AddMsg(.hWnd, WM_CANCELMODE)
         Call Subclass_AddMsg(.hWnd, WM_LBUTTONDBLCLK)
         Call Subclass_AddMsg(.hWnd, WM_LBUTTONDOWN)
         Call Subclass_AddMsg(.hWnd, WM_LBUTTONUP)
         Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE)
         Call Subclass_AddMsg(.hWnd, WM_MOUSEWHEEL)
         Call Subclass_AddMsg(.hWnd, WM_PAINT, MSG_BEFORE)
         Call Subclass_AddMsg(.hWnd, WM_SIZE, MSG_BEFORE)
         Call Subclass_AddMsg(.hWnd, WM_SYSCOLORCHANGE)
         Call Subclass_AddMsg(.hWnd, WM_TIMER)
         Call Subclass_Initialize(ContainerHwnd)
         Call Subclass_AddMsg(ContainerHwnd, WM_KEYDOWN, MSG_BEFORE)
         Call Subclass_AddMsg(ContainerHwnd, WM_KEYUP, MSG_BEFORE)
         Call Subclass_AddMsg(ContainerHwnd, WM_MOUSEWHEEL)
         
         If IsThemedWindows Then Call Subclass_AddMsg(.hWnd, WM_THEMECHANGED)
      End With
   End If

End Sub

Private Sub UserControl_Resize()

   If InitProperties Then
      If Width > Height Then
         m_Orientation = Horizontal
         
      Else
         m_Orientation = Vertical
      End If
      
      InitProperties = False
   End If
   
   If Not Ambient.UserMode Then Call WhenSize

End Sub

Private Sub UserControl_Terminate()

   On Local Error GoTo ExitSub
   
   Call Subclass_Terminate
   
ExitSub:
   On Local Error GoTo 0
   
   Call TimerKill(TIMERID_HOT)
   Call TimerKill(TIMERID_CHANGE1)
   Call TimerKill(TIMERID_CHANGE2)
   
   DeleteObject hPatternBrush
   Erase SubclassData

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "ContainerArrowKeys", m_ContainerArrowKeys, False
      .WriteProperty "Enabled", UserControl.Enabled, True
      .WriteProperty "LargeChange", m_LargeChange, 1
      .WriteProperty "Max", m_Max, 0
      .WriteProperty "Min", m_Min, 0
      .WriteProperty "MouseWheel", m_MouseWheel, True
      .WriteProperty "MouseWheelInContainer", m_MouseWheelInContainer, False
      .WriteProperty "ShowButtons", m_ShowButtons, True
      .WriteProperty "SmallChange", m_SmallChange, 1
      .WriteProperty "Orientation", m_Orientation, Vertical
      .WriteProperty "Value", m_Value, 0
   End With

End Sub
