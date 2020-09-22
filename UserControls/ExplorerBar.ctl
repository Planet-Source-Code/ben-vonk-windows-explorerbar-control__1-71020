VERSION 5.00
Begin VB.UserControl ExplorerBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2316
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1812
   KeyPreview      =   -1  'True
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   151
   ToolboxBitmap   =   "ExplorerBar.ctx":0000
   Begin prjExplorerBar.Border bdrBorder 
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   445
      BorderColor     =   -2147483640
   End
   Begin VB.PictureBox picGroupMasker 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   840
      Picture         =   "ExplorerBar.ctx":0312
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   168
   End
   Begin VB.PictureBox picButtons 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   324
      Left            =   840
      Picture         =   "ExplorerBar.ctx":06F0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   168
   End
   Begin prjExplorerBar.ThemedScrollBar tsbVertical 
      Height          =   1572
      Left            =   1200
      TabIndex        =   5
      Top             =   0
      Width           =   252
      _ExtentX        =   445
      _ExtentY        =   2773
      LargeChange     =   50
      MouseWheelInContainer=   -1  'True
      SmallChange     =   10
   End
   Begin VB.PictureBox picAnimation 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox picGroup 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   252
      Index           =   0
      Left            =   120
      MouseIcon       =   "ExplorerBar.ctx":0BD6
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   1200
   End
   Begin VB.PictureBox picItems 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   492
      Index           =   0
      Left            =   120
      MouseIcon       =   "ExplorerBar.ctx":0EE0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox picButtonMasker 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   228
      Left            =   840
      Picture         =   "ExplorerBar.ctx":11EA
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   228
   End
End
Attribute VB_Name = "ExplorerBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ExplorerBar Control
'
'Author Ben Vonk
'08-03-2008 First version (Based on Alex Flex's 'XP-Style ExplorerBar control - version 2' at http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=61899&lngWId=1)
'09-10-2008 Second version, Add ThemedScrollBar second version and fixed some minor bugs
'28-12-2011 Third version, Fixed some bugs, add more functions and features

Option Explicit

' Public Events
Public Event Collapse(Group As Integer)
Public Event Expand(Group As Integer)
Public Event ErrorOpenFile(Group As Integer, Item As Integer, File As String, Error As Long)
Public Event GroupClick(Group As Integer, WindowState As WindowStates)
Public Event ItemClick(Group As Integer, Item As Integer)
Public Event ItemOpenFile(Group As Integer, Item As Integer, File As String)
Public Event MouseDown(Group As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseHover(Group As Integer, Item As Integer, FullTextShowed As Boolean)
Public Event MouseMove(Group As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseOut(Group As Integer, Item As Integer)
Public Event MouseUp(Group As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

' Private Constants
Private Const EBP_HEADERBACKGROUND       As Integer = 1
Private Const EBP_NORMALGROUPBACKGROUND  As Integer = 5
Private Const GROUP_SPACE                As Integer = 15
Private Const ITEM_SPACE                 As Integer = 10
Private Const EBP_NORMALGROUPHEAD        As Long = 8
Private Const EBP_SPECIALGROUPHEAD       As Long = 12
Private Const ALL_MESSAGES               As Long = -1
Private Const DT_LEFT                    As Long = &H0
Private Const DT_WORD_ELLIPSIS           As Long = &H40000
Private Const GROUP_DETAILS              As Long = 3
Private Const GROUP_NORMAL               As Long = 1
Private Const GROUP_SPECIAL              As Long = 2
Private Const GWL_WNDPROC                As Long = -4
Private Const PATCH_05                   As Long = 93
Private Const PATCH_09                   As Long = 137
Private Const STATE_HOT                  As Long = 2
Private Const STATE_NORMAL               As Long = 1
Private Const STATE_PRESSED              As Long = 3

Private Const WM_LBUTTONDOWN             As Long = &H201
Private Const WM_MOUSELEAVE              As Long = &H2A3
Private Const WM_MOUSEMOVE               As Long = &H200
Private Const WM_MOUSEWHEEL              As Long = &H20A
Private Const WM_SYSCOLORCHANGE          As Long = &H15
Private Const WM_THEMECHANGED            As Long = &H31A
Private Const THEME_EMBEDDED             As String = "Embedded"
Private Const THEME_LUNA                 As String = "Luna"
Private Const THEME_MEDIACENTRE          As String = "Media Centre"
Private Const THEME_NAME                 As String = "ExplorerBar"
Private Const THEME_ROYALE               As String = "Royale"
Private Const THEME_ZUNE                 As String = "Zune"
Private Const VERSION                    As String = "3.00"

' Public Enumaration
Public Enum Animations
   None
   Slow
   Medium
   Fast
End Enum

Public Enum ButtonHeights
   Low
   High
End Enum

' Private Enumerations
Private Enum ColorsRGB
   IsRed
   IsGreen
   IsBlue
End Enum

Public Enum ExploreBarThemeTypes
   Windows
   Enhanced
   User
End Enum

Public Enum GradientStyles
   TopBottom
   BottomTop
   LeftRight
   RightLeft
End Enum

Private Enum GroupItemProperties
   IsDetailCaption
   IsDetailPicture
   IsDetailTitle
   IsGroupIcon
   IsGroupContainer
   IsGroupItemsBackgroundPicture
   IsGroupTag
   IsGroupTitle
   IsGroupTitleBold
   IsItemCaption
   IsItemCaptionBold
   IsItemOpenFile
   IsItemTextOnly
   IsItemIcon
   IsItemTag
   IsToolTipText
   IsWindowState
End Enum

Private Enum MsgWhen
   MSG_BEFORE = 1
   MSG_AFTER = 2
   MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
End Enum

Public Enum WindowStates
   Expanded
   Collapsed
   Fixed
End Enum

' Private Types
Private Type BlendFunctionType
   BlendOp                               As Byte
   BlendFlags                            As Byte
   SourceConstantAlpha                   As Byte
   AlphaFormat                           As Byte
End Type

Private Type GradientRect
   UpperLeft                             As Long
   LowerRight                            As Long
End Type

Private Type Rect
   Left                                  As Long
   Top                                   As Long
   Right                                 As Long
   Bottom                                As Long
End Type

Private Type GroupItemType
   State                                 As Long
   Rect                                  As Rect
   Caption                               As String
   Bold                                  As Boolean
   FullTextShowed                        As Boolean
   TextOnly                              As Boolean
   OpenFile                              As String
   Tag                                   As String
   ToolTipText                           As String
   Icon                                  As IPictureDisp
End Type

Private Type GroupType
   TypeGroup                             As Long
   State                                 As Long
   WindowState                           As WindowStates
   Icon                                  As IPictureDisp
   Title                                 As String
   Bold                                  As Boolean
   FullTextShowed                        As Boolean
   Tag                                   As String
   ToolTipText                           As String
   Items()                               As GroupItemType
   ItemsBackgroundPicture                As IPictureDisp
   ItemsCount                            As Integer
   DetailPicture                         As IPictureDisp
   DetailTitle                           As String
   DetailCaption                         As String
   Container                             As PictureBox
   ContainerScaleMode                    As Integer
   ContainerSet                          As Boolean
End Type

Private Type PointAPI
   X                                     As Long
   Y                                     As Long
End Type

Private Type SubclassDataType
   hWnd                                  As Long
   nAddrSclass                           As Long
   nAddrOrig                             As Long
   nMsgCountA                            As Long
   nMsgCountB                            As Long
   aMsgTabelA()                          As Long
   aMsgTabelB()                          As Long
End Type

Private Type TrackMouseEventStruct
   cbSize                                As Long
   dwFlags                               As Long
   hwndTrack                             As Long
   dwHoverTime                           As Long
End Type

Private Type TriVertex
   X                                     As Long
   Y                                     As Long
   Red                                   As Integer
   Green                                 As Integer
   Blue                                  As Integer
   Alpha                                 As Integer
End Type

' Private Variables
Private m_Animation                      As Animations
Private DoAlignment                      As Boolean
Private DoRoundGroup                     As Boolean
Private FreezeMouseMove                  As Boolean
Private GroupDetailsSet                  As Boolean
Private GroupsExist                      As Boolean
Private GroupSpecialsSet                 As Boolean
Private HasTheme                         As Boolean
Private IsMouseDown                      As Boolean
Private KeyPressed                       As Boolean
Private MouseHoverControl                As Boolean
Private MouseHoverScrollBar              As Boolean
Private m_DetailGroupButton              As Boolean
Private m_Locked                         As Boolean
Private m_OpenOneGroupOnly               As Boolean
Private m_SoundGroupClicked              As Boolean
Private m_SoundItemClicked               As Boolean
Private m_UseAlphaBlend                  As Boolean
Private m_UseUserForeColors              As Boolean
Private StateNormalGroup                 As Boolean
Private StateSpecialGroup                As Boolean
Private TrackUser32                      As Boolean
Private UseButtonPictures                As Boolean
Private m_HeaderHeight                   As ButtonHeights
Private SubclassFunk                     As Collection
Private m_UseTheme                       As ExploreBarThemeTypes
Private m_GradientStyle                  As GradientStyles
Private Groups()                         As GroupType
Private AlphaPercent                     As Integer
Private BlendStep                        As Integer
Private ClickedItem                      As Integer
Private FocussedGroup                    As Integer
Private FocussedItem                     As Integer
Private GroupAnimated                    As Integer
Private GroupsCount                      As Integer
Private GroupIndex                       As Integer
Private MouseHoverGroup                  As Integer
Private MouseHoverItem                   As Integer
Private MoveLines                        As Integer
Private PressedGroup                     As Integer
Private PressedItem                      As Integer
Private m_ButtonPicture(11)              As IPictureDisp
Private BorderColorItems                 As Long
Private ForeColorHoverGroups             As Long
Private ForeColorHoverGroupSpecial       As Long
Private ForeColorHoverItems              As Long
Private ForeColorNormalGroups            As Long
Private ForeColorNormalGroupSpecial      As Long
Private ForeColorNormalItems             As Long
Private ItemHeight                       As Long
Private m_BackColor                      As Long
Private m_DetailsForeColor               As Long
Private m_GradientBackColor              As Long
Private m_GradientNormalHeaderBackColor  As Long
Private m_GradientNormalItemBackColor    As Long
Private m_GradientSpecialHeaderBackColor As Long
Private m_GradientSpecialItemBackColor   As Long
Private m_NormalArrowDownColor           As Long
Private m_NormalArrowHoverColor          As Long
Private m_NormalArrowUpColor             As Long
Private m_NormalButtonBackColor          As Long
Private m_NormalButtonDownColor          As Long
Private m_NormalButtonHoverColor         As Long
Private m_NormalButtonUpColor            As Long
Private m_NormalHeaderBackColor          As Long
Private m_NormalHeaderForeColor          As Long
Private m_NormalHeaderHoverColor         As Long
Private m_NormalItemBackColor            As Long
Private m_NormalItemBorderColor          As Long
Private m_NormalItemForeColor            As Long
Private m_NormalItemHoverColor           As Long
Private m_SpecialArrowDownColor          As Long
Private m_SpecialArrowHoverColor         As Long
Private m_SpecialArrowUpColor            As Long
Private m_SpecialButtonBackColor         As Long
Private m_SpecialButtonDownColor         As Long
Private m_SpecialButtonHoverColor        As Long
Private m_SpecialButtonUpColor           As Long
Private m_SpecialHeaderBackColor         As Long
Private m_SpecialHeaderForeColor         As Long
Private m_SpecialHeaderHoverColor        As Long
Private m_SpecialItemBackColor           As Long
Private m_SpecialItemBorderColor         As Long
Private m_SpecialItemForeColor           As Long
Private m_SpecialItemHoverColor          As Long
Private ScrollLines                      As Long
Private ThemeMap                         As String
Private ThemeName                        As String
Private ThemeColorMap                    As String
Private ThemeColorName                   As String
Private SubclassData()                   As SubclassDataType

' Private API's
Private Declare Function TrackMouseEventComCtl Lib "ComCtl32" Alias "_TrackMouseEvent" (lpEventTrack As TrackMouseEventStruct) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCdest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCsrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Integer
Private Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function StrLen Lib "Kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function AlphaBlend Lib "MSImg32" (ByVal hDCdest As Long, ByVal Xdest As Long, ByVal Ydest As Long, ByVal Widthdest As Long, ByVal Heightdest As Long, ByVal hDCsrc As Long, ByVal Xsrc As Long, ByVal Ysrc As Long, ByVal Widthsrc As Long, ByVal Heightsrc As Long, ByVal BlendFunction As Long) As Long
Private Declare Function GradientFill Lib "MSImg32" (ByVal hDC As Long, ByRef pVertex As TriVertex, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Integer
Private Declare Function TransparentBlt Lib "MSImg32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal Xsrc As Long, ByVal Ysrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function SafeArrayGetDim Lib "OleAut32" (ByRef saArray() As Any) As Long
Private Declare Function OleTranslateColor Lib "OLEPro32" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function ShellExecute Lib "Shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function DrawFocusRect Lib "User32" (ByVal hDC As Long, ByRef lpRect As Rect) As Long
Private Declare Function DrawText Lib "User32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function PtInRect Lib "User32" (Rect As Rect, ByVal lPtX As Long, ByVal lPtY As Long) As Integer
Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLongA Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function TrackMouseEvent Lib "User32" (lpEventTrack As TrackMouseEventStruct) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function CloseThemeData Lib "UxTheme" (ByVal lngTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "UxTheme" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, rctFrame As Rect, rctClip As Rect) As Long
Private Declare Function GetCurrentThemeName Lib "UxTheme" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeColor Lib "UxTheme" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, pColor As Long) As Long
Private Declare Function OpenThemeData Lib "UxTheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function SoundPlay Lib "WinMM" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Subclass_WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lhWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

Dim lngWindow As Long

   If uMsg = WM_MOUSELEAVE Then
      lngWindow = CheckMouseInControl
      MouseHoverGroup = -1
      MouseHoverItem = -1
      
      If lhWnd = tsbVertical.hWnd Then
         MouseHoverScrollBar = False
         RaiseEvent MouseOut(-2, -1)
         
         If lngWindow <> UserControl.hWnd Then
            MouseHoverControl = False
            RaiseEvent MouseOut(-1, -1)
         End If
      End If
      
      If (lhWnd = UserControl.hWnd) And (lngWindow <> UserControl.hWnd) Then
         If CheckMouseInContainer(lhWnd, lngWindow) Then Exit Sub
         
         MouseHoverControl = False
         RaiseEvent MouseOut(-1, -1)
      End If
      
   ElseIf uMsg = WM_MOUSEMOVE Then
      Call TrackMouseLeave(lhWnd)
      
      If (MouseHoverGroup > -1) And Not KeyPressed Then
         RaiseEvent MouseOut(MouseHoverGroup, -1)
         MouseHoverGroup = -1
         MouseHoverItem = -1
      End If
      
      If Not MouseHoverControl Then
         MouseHoverControl = True
         ClickedItem = -1
         
         If (lhWnd = UserControl.hWnd) Or (lhWnd = tsbVertical.hWnd) Then RaiseEvent MouseHover(-1, -1, False)
      End If
      
   ElseIf uMsg = WM_MOUSEWHEEL Then
      Call tsbVertical_MouseWheel(ScrollLines + ((ScrollLines * 2) * (wParam > 0)))
      
   ElseIf (uMsg = WM_THEMECHANGED) Or (uMsg = WM_SYSCOLORCHANGE) Then
      DoEvents
      HasTheme = False
      
      Call GetTextColors
      Call Refresh
      
   ElseIf uMsg = WM_LBUTTONDOWN Then
      If FocussedGroup > -1 Then Call ResetFocussedGroup
   End If

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

Private Sub Subclass_DelMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As MsgWhen = MSG_AFTER)

   With SubclassData(Subclass_Index(lhWnd))
      If When And MSG_BEFORE Then Call Subclass_DoDelMsg(uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass)
      If When And MSG_AFTER Then Call Subclass_DoDelMsg(uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass)
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

Private Sub Subclass_DoDelMsg(ByVal uMsg As Long, ByRef aMsgTabel() As Long, ByRef nMsgCount As Long, ByVal When As MsgWhen, ByVal nAddr As Long)

Dim lngEntry As Long

   If uMsg = ALL_MESSAGES Then
      nMsgCount = 0
      
      If When = MSG_BEFORE Then
         lngEntry = PATCH_05
         
      Else
         lngEntry = PATCH_09
      End If
      
      Call Subclass_PatchVal(nAddr, lngEntry, 0)
      
   Else
      For lngEntry = 1 To nMsgCount - 1
         If aMsgTabel(lngEntry) = uMsg Then
            aMsgTabel(lngEntry) = 0
            Exit For
         End If
      Next 'lngEntry
   End If

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

Public Property Get Animation() As Animations
Attribute Animation.VB_Description = "Returns/sets a the speed that will be used for the open and close window animations."

   Animation = m_Animation

End Property

Public Property Let Animation(ByVal NewAnimation As Animations)

   m_Animation = NewAnimation
   PropertyChanged "Animation"

End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphcs in an object."

   BackColor = m_BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

   m_BackColor = NewBackColor
   PropertyChanged "BackColor"
   
   Call Refresh

End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the border color used to display the border in an object."

   BorderColor = bdrBorder.BorderColor

End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)

   bdrBorder.BorderColor = NewBorderColor
   PropertyChanged "BorderColor"

End Property

Public Property Get DetailsForeColor() As OLE_COLOR
Attribute DetailsForeColor.VB_Description = "Returns/sets the forground color used to display text in an detail window."

   DetailsForeColor = m_DetailsForeColor

End Property

Public Property Let DetailsForeColor(ByVal NewDetailsForeColor As OLE_COLOR)

   m_DetailsForeColor = NewDetailsForeColor
   PropertyChanged "DetailsForeColor"
   
   Call Refresh

End Property

Public Property Get DetailGroupButton() As Boolean
Attribute DetailGroupButton.VB_Description = "Determins whether the DetailGroup has a button."

   DetailGroupButton = m_DetailGroupButton

End Property

Public Property Let DetailGroupButton(ByVal NewDetailsGroupHasButton As Boolean)

Dim intGroup As Integer

   If GroupsExist Then
      For intGroup = 0 To UBound(Groups)
         With Groups(intGroup)
            If .TypeGroup = GROUP_DETAILS Then
               If .WindowState = Fixed Then
                  If NewDetailsGroupHasButton Then .WindowState = Expanded
                  
               ElseIf Not NewDetailsGroupHasButton Then
                  If .WindowState = Collapsed Then Call picGroup_MouseUp(intGroup, vbLeftButton, 0, 0, 0)
                  
                  .WindowState = Fixed
               End If
            End If
         End With
      Next 'intGroup
   End If
   
   m_DetailGroupButton = NewDetailsGroupHasButton
   PropertyChanged "DetailGroupButton"

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."

   Set Font = UserControl.Font

End Property

Public Property Let Font(ByRef NewFont As StdFont)

   Set Font = NewFont

End Property

Public Property Set Font(ByRef NewFont As StdFont)

Dim intGroup As Integer

   With NewFont
      If .Size < 8 Then .Size = 8
      If .Size > 10 Then .Size = 10
   End With
   
   UserControl.Font = NewFont
   UserControl.FontBold = False
   PropertyChanged "Font"
   
   For intGroup = 0 To picGroup.Count - 1
      picGroup.Item(intGroup).Font = UserControl.Font
   Next 'intGroup
   
   For intGroup = 0 To picItems.Count - 1
      picItems.Item(intGroup).Font = UserControl.Font
   Next 'intGroup
   
   Call Refresh

End Property

Public Property Get GradientBackColor() As OLE_COLOR
Attribute GradientBackColor.VB_Description = "Returns/sets the gradient color for an object background."

   GradientBackColor = m_GradientBackColor

End Property

Public Property Let GradientBackColor(ByVal NewGradientBackColor As OLE_COLOR)

   m_GradientBackColor = NewGradientBackColor
   PropertyChanged "GradientBackColor"
   
   Call Refresh

End Property

Public Property Get GradientNormalHeaderBackColor() As OLE_COLOR
Attribute GradientNormalHeaderBackColor.VB_Description = "Returns/sets the gradient color for an normal header background."

   GradientNormalHeaderBackColor = m_GradientNormalHeaderBackColor

End Property

Public Property Let GradientNormalHeaderBackColor(ByVal NewGradientNormalHeaderBackColor As OLE_COLOR)

   m_GradientNormalHeaderBackColor = NewGradientNormalHeaderBackColor
   PropertyChanged "GradientNormalHeaderBackColor"
   
   Call Refresh

End Property

Public Property Get GradientNormalItemBackColor() As OLE_COLOR
Attribute GradientNormalItemBackColor.VB_Description = "Returns/sets the gradient color for an normal item background."

   GradientNormalItemBackColor = m_GradientNormalItemBackColor

End Property

Public Property Let GradientNormalItemBackColor(ByVal NewGradientNormalItemBackColor As OLE_COLOR)

   m_GradientNormalItemBackColor = NewGradientNormalItemBackColor
   PropertyChanged "GradientNormalItemBackColor"
   
   Call Refresh

End Property

Public Property Get GradientSpecialHeaderBackColor() As OLE_COLOR
Attribute GradientSpecialHeaderBackColor.VB_Description = "Returns/sets the gradient color for an special header background."

   GradientSpecialHeaderBackColor = m_GradientSpecialHeaderBackColor

End Property

Public Property Let GradientSpecialHeaderBackColor(ByVal NewGradientSpecialHeaderBackColor As OLE_COLOR)

   m_GradientSpecialHeaderBackColor = NewGradientSpecialHeaderBackColor
   PropertyChanged "GradientSpecialHeaderBackColor"
   
   Call Refresh

End Property

Public Property Get GradientSpecialItemBackColor() As OLE_COLOR
Attribute GradientSpecialItemBackColor.VB_Description = "Returns/sets the gradient color for an special item background."

   GradientSpecialItemBackColor = m_GradientSpecialItemBackColor

End Property

Public Property Let GradientSpecialItemBackColor(ByVal NewGradientSpecialItemBackColor As OLE_COLOR)

   m_GradientSpecialItemBackColor = NewGradientSpecialItemBackColor
   PropertyChanged "GradientSpecialItemBackColor"
   
   Call Refresh

End Property

Public Property Get GradientStyle() As GradientStyles
Attribute GradientStyle.VB_Description = "Returns/sets the style to draw the gradient color."

   GradientStyle = m_GradientStyle

End Property

Public Property Let GradientStyle(ByVal NewGradientStyle As GradientStyles)

   m_GradientStyle = NewGradientStyle
   PropertyChanged "GradientStyle"
   
   Call Refresh

End Property

Public Property Get HeaderHeight() As ButtonHeights
Attribute HeaderHeight.VB_Description = "Returns/sets the height of the group headers."

   HeaderHeight = m_HeaderHeight

End Property

Public Property Let HeaderHeight(ByVal NewHeaderHeight As ButtonHeights)

   m_HeaderHeight = NewHeaderHeight
   PropertyChanged "HeaderHeight"
   
   Call Refresh

End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets to lock/unlock an object to fasten the property changes."

   Locked = m_Locked

End Property

Public Property Let Locked(ByVal NewLocked As Boolean)

   If m_Locked = NewLocked Then Exit Property
   
   m_Locked = NewLocked
   PropertyChanged "Locked"
   
   If Not m_Locked Then Call Refresh

End Property

Public Property Get NormalArrowDownColor() As OLE_COLOR
Attribute NormalArrowDownColor.VB_Description = "Returns/sets the arrow color when an normal header is down."

   NormalArrowDownColor = m_NormalArrowDownColor

End Property

Public Property Let NormalArrowDownColor(ByVal NewNormalArrowDownColor As OLE_COLOR)

   m_NormalArrowDownColor = NewNormalArrowDownColor
   PropertyChanged "NormalArrowDownColor"
   
   Call Refresh

End Property

Public Property Get NormalArrowHoverColor() As OLE_COLOR
Attribute NormalArrowHoverColor.VB_Description = "Returns/sets the arrow color when the mouse is hover an normal header."

   NormalArrowHoverColor = m_NormalArrowHoverColor

End Property

Public Property Let NormalArrowHoverColor(ByVal NewNormalArrowHoverColor As OLE_COLOR)

   m_NormalArrowHoverColor = NewNormalArrowHoverColor
   PropertyChanged "NormalArrowHoverColor"
   
   Call Refresh

End Property

Public Property Get NormalArrowUpColor() As OLE_COLOR
Attribute NormalArrowUpColor.VB_Description = "Returns/sets the arrow color when an normal header is up."

   NormalArrowUpColor = m_NormalArrowUpColor

End Property

Public Property Let NormalArrowUpColor(ByVal NewNormalArrowUpColor As OLE_COLOR)

   m_NormalArrowUpColor = NewNormalArrowUpColor
   PropertyChanged "NormalArrowUpColor"
   
   Call Refresh

End Property

Public Property Get NormalButtonBackColor() As OLE_COLOR
Attribute NormalButtonBackColor.VB_Description = "Returns/sets the background color for an normal header button."

   NormalButtonBackColor = m_NormalButtonBackColor

End Property

Public Property Let NormalButtonBackColor(ByVal NewNormalButtonBackColor As OLE_COLOR)

   m_NormalButtonBackColor = NewNormalButtonBackColor
   PropertyChanged "NormalButtonBackColor"
   
   Call Refresh

End Property

Public Property Get NormalButtonDownColor() As OLE_COLOR
Attribute NormalButtonDownColor.VB_Description = "Returns/sets the color for an normal header button when it is down."

   NormalButtonDownColor = m_NormalButtonDownColor

End Property

Public Property Let NormalButtonDownColor(ByVal NewNormalButtonDownColor As OLE_COLOR)

   m_NormalButtonDownColor = NewNormalButtonDownColor
   PropertyChanged "NormalButtonDownColor"
   
   Call Refresh

End Property

Public Property Get NormalButtonHoverColor() As OLE_COLOR
Attribute NormalButtonHoverColor.VB_Description = "Returns/sets the color for an normal header button when the mouse is hover the header."

   NormalButtonHoverColor = m_NormalButtonHoverColor

End Property

Public Property Let NormalButtonHoverColor(ByVal NewNormalButtonHoverColor As OLE_COLOR)

   m_NormalButtonHoverColor = NewNormalButtonHoverColor
   PropertyChanged "NormalButtonHoverColor"
   
   Call Refresh

End Property

Public Property Get NormalButtonPictureDown() As IPictureDisp
Attribute NormalButtonPictureDown.VB_Description = "Returns/sets a graphic for an normal header button when it is down."

   Set NormalButtonPictureDown = m_ButtonPicture(2)

End Property

Public Property Let NormalButtonPictureDown(ByVal NewNormalButtonPictureDown As IPictureDisp)

   Set NormalButtonPictureDown = NewNormalButtonPictureDown

End Property

Public Property Set NormalButtonPictureDown(ByVal NewNormalButtonPictureDown As IPictureDisp)

   Call SetButtonPicture(2, NewNormalButtonPictureDown, "NormalButtonPictureDown", False)

End Property

Public Property Get NormalButtonPictureHover() As IPictureDisp
Attribute NormalButtonPictureHover.VB_Description = "Returns/sets a graphic for an normal header button when the mouse is hover the header."

   Set NormalButtonPictureHover = m_ButtonPicture(1)

End Property

Public Property Let NormalButtonPictureHover(ByVal NewNormalButtonPictureHover As IPictureDisp)

   Set NormalButtonPictureHover = NewNormalButtonPictureHover

End Property

Public Property Set NormalButtonPictureHover(ByVal NewNormalButtonPictureHover As IPictureDisp)

   Call SetButtonPicture(1, NewNormalButtonPictureHover, "NormalButtonPictureHover", False)

End Property

Public Property Get NormalButtonPictureUp() As IPictureDisp
Attribute NormalButtonPictureUp.VB_Description = "Returns/sets a graphic for an normal header button when it is up."

   Set NormalButtonPictureUp = m_ButtonPicture(0)

End Property

Public Property Let NormalButtonPictureUp(ByVal NewNormalButtonPictureUp As IPictureDisp)

   Set NormalButtonPictureUp = NewNormalButtonPictureUp

End Property

Public Property Set NormalButtonPictureUp(ByVal NewNormalButtonPictureUp As IPictureDisp)

   Call SetButtonPicture(0, NewNormalButtonPictureUp, "NormalButtonPictureUp", False)

End Property

Public Property Get NormalButtonUpColor() As OLE_COLOR
Attribute NormalButtonUpColor.VB_Description = "Returns/sets the color for an normal header button when it is up."

   NormalButtonUpColor = m_NormalButtonUpColor

End Property

Public Property Let NormalButtonUpColor(ByVal NewNormalButtonUpColor As OLE_COLOR)

   m_NormalButtonUpColor = NewNormalButtonUpColor
   PropertyChanged "NormalButtonUpColor"
   
   Call Refresh

End Property

Public Property Get NormalHeaderBackColor() As OLE_COLOR
Attribute NormalHeaderBackColor.VB_Description = "Returns/sets the background color for an normal header."

   NormalHeaderBackColor = m_NormalHeaderBackColor

End Property

Public Property Let NormalHeaderBackColor(ByVal NewNormalHeaderBackColor As OLE_COLOR)

   m_NormalHeaderBackColor = NewNormalHeaderBackColor
   PropertyChanged "NormalHeaderBackColor"
   
   Call Refresh

End Property

Public Property Get NormalHeaderForeColor() As OLE_COLOR
Attribute NormalHeaderForeColor.VB_Description = "Returns/sets the foreground color for an normal header."

   NormalHeaderForeColor = m_NormalHeaderForeColor

End Property

Public Property Let NormalHeaderForeColor(ByVal NewNormalHeaderForeColor As OLE_COLOR)

   m_NormalHeaderForeColor = NewNormalHeaderForeColor
   PropertyChanged "NormalHeaderForeColor"
   
   Call Refresh

End Property

Public Property Get NormalHeaderHoverColor() As OLE_COLOR
Attribute NormalHeaderHoverColor.VB_Description = "Returns/sets the color for an normal header when the mouse is hover the header."

   NormalHeaderHoverColor = m_NormalHeaderHoverColor

End Property

Public Property Let NormalHeaderHoverColor(ByVal NewNormalHeaderHoverColor As OLE_COLOR)

   m_NormalHeaderHoverColor = NewNormalHeaderHoverColor
   PropertyChanged "NormalHeaderHoverColor"
   
   Call Refresh

End Property

Public Property Get NormalItemBackColor() As OLE_COLOR
Attribute NormalItemBackColor.VB_Description = "Returns/sets the background color for an normal item group."

   NormalItemBackColor = m_NormalItemBackColor

End Property

Public Property Let NormalItemBackColor(ByVal NewNormalItemBackColor As OLE_COLOR)

   m_NormalItemBackColor = NewNormalItemBackColor
   PropertyChanged "NormalItemBackColor"
   
   Call Refresh

End Property

Public Property Get NormalItemBorderColor() As OLE_COLOR
Attribute NormalItemBorderColor.VB_Description = "Returns/sets the border color for an normal item group."

   NormalItemBorderColor = m_NormalItemBorderColor

End Property

Public Property Let NormalItemBorderColor(ByVal NewNormalItemBorderColor As OLE_COLOR)

   m_NormalItemBorderColor = NewNormalItemBorderColor
   PropertyChanged "NormalItemBorderColor"
   
   Call Refresh

End Property

Public Property Get NormalItemForeColor() As OLE_COLOR
Attribute NormalItemForeColor.VB_Description = "Returns/sets the foreground color for an normal item group."

   NormalItemForeColor = m_NormalItemForeColor

End Property

Public Property Let NormalItemForeColor(ByVal NewNormalItemForeColor As OLE_COLOR)

   m_NormalItemForeColor = NewNormalItemForeColor
   PropertyChanged "NormalItemForeColor"
   
   Call Refresh

End Property

Public Property Get NormalItemHoverColor() As OLE_COLOR
Attribute NormalItemHoverColor.VB_Description = "Returns/sets the color for an normal item group when the mouse is hover the item."

   NormalItemHoverColor = m_NormalItemHoverColor

End Property

Public Property Let NormalItemHoverColor(ByVal NewNormalItemHoverColor As OLE_COLOR)

   m_NormalItemHoverColor = NewNormalItemHoverColor
   PropertyChanged "NormalItemHoverColor"
   
   Call Refresh

End Property

Public Property Get OpenOneGroupOnly() As Boolean
Attribute OpenOneGroupOnly.VB_Description = "Determins whether one groups is open at a time, except fixed groups."

   OpenOneGroupOnly = m_OpenOneGroupOnly

End Property

Public Property Let OpenOneGroupOnly(ByVal NewOpenOneGroupOnly As Boolean)

   m_OpenOneGroupOnly = NewOpenOneGroupOnly
   PropertyChanged "OpenOneGroupOnly"
   
   Call Refresh

End Property

Public Property Get ShowBorder() As Boolean
Attribute ShowBorder.VB_Description = "Determins whether the object border will be showed."

   ShowBorder = bdrBorder.Visible

End Property

Public Property Let ShowBorder(ByVal NewShowBorder As Boolean)

   bdrBorder.Visible = NewShowBorder

End Property

Public Property Get SoundGroupClicked() As Boolean
Attribute SoundGroupClicked.VB_Description = "Determins whether a sound will be played when a group is clicked."

   SoundGroupClicked = m_SoundGroupClicked

End Property

Public Property Let SoundGroupClicked(ByVal NewSoundGroupClicked As Boolean)

   m_SoundGroupClicked = NewSoundGroupClicked
   PropertyChanged "SoundGroupClicked"

End Property

Public Property Get SoundItemClicked() As Boolean
Attribute SoundItemClicked.VB_Description = "Determins whether a sound will be played when a item is clicked."

   SoundItemClicked = m_SoundItemClicked

End Property

Public Property Let SoundItemClicked(ByVal NewSoundItemClicked As Boolean)

   m_SoundItemClicked = NewSoundItemClicked
   PropertyChanged "SoundItemClicked"

End Property

Public Property Get SpecialArrowDownColor() As OLE_COLOR
Attribute SpecialArrowDownColor.VB_Description = "Returns/sets the arrow color when an special header is up."

   SpecialArrowDownColor = m_SpecialArrowDownColor

End Property

Public Property Let SpecialArrowDownColor(ByVal NewSpecialArrowDownColor As OLE_COLOR)

   m_SpecialArrowDownColor = NewSpecialArrowDownColor
   PropertyChanged "SpecialArrowDownColor"
   
   Call Refresh

End Property

Public Property Get SpecialArrowHoverColor() As OLE_COLOR
Attribute SpecialArrowHoverColor.VB_Description = "Returns/sets the arrow color when the mouse is hover an special header."

   SpecialArrowHoverColor = m_SpecialArrowHoverColor

End Property

Public Property Let SpecialArrowHoverColor(ByVal NewSpecialArrowHoverColor As OLE_COLOR)

   m_SpecialArrowHoverColor = NewSpecialArrowHoverColor
   PropertyChanged "SpecialArrowHoverColor"
   
   Call Refresh

End Property

Public Property Get SpecialArrowUpColor() As OLE_COLOR
Attribute SpecialArrowUpColor.VB_Description = "Returns/sets the arrow color when an special header is up."

   SpecialArrowUpColor = m_SpecialArrowUpColor

End Property

Public Property Let SpecialArrowUpColor(ByVal NewSpecialArrowUpColor As OLE_COLOR)

   m_SpecialArrowUpColor = NewSpecialArrowUpColor
   PropertyChanged "SpecialArrowUpColor"
   
   Call Refresh

End Property

Public Property Get SpecialButtonBackColor() As OLE_COLOR
Attribute SpecialButtonBackColor.VB_Description = "Returns/sets the background color for an special header button."

   SpecialButtonBackColor = m_SpecialButtonBackColor

End Property

Public Property Let SpecialButtonBackColor(ByVal NewSpecialButtonBackColor As OLE_COLOR)

   m_SpecialButtonBackColor = NewSpecialButtonBackColor
   PropertyChanged "SpecialButtonBackColor"
   
   Call Refresh

End Property

Public Property Get SpecialButtonDownColor() As OLE_COLOR
Attribute SpecialButtonDownColor.VB_Description = "Returns/sets the color for an special header button when it is down."

   SpecialButtonDownColor = m_SpecialButtonDownColor

End Property

Public Property Let SpecialButtonDownColor(ByVal NewSpecialButtonDownColor As OLE_COLOR)

   m_SpecialButtonDownColor = NewSpecialButtonDownColor
   PropertyChanged "SpecialButtonDownColor"
   
   Call Refresh

End Property

Public Property Get SpecialButtonHoverColor() As OLE_COLOR
Attribute SpecialButtonHoverColor.VB_Description = "Returns/sets the color for an special header button when the mouse is hover the header."

   SpecialButtonHoverColor = m_SpecialButtonHoverColor

End Property

Public Property Let SpecialButtonHoverColor(ByVal NewSpecialButtonHoverColor As OLE_COLOR)

   m_SpecialButtonHoverColor = NewSpecialButtonHoverColor
   PropertyChanged "SpecialButtonHoverColor"
   
   Call Refresh

End Property

Public Property Get SpecialButtonPictureDown() As IPictureDisp
Attribute SpecialButtonPictureDown.VB_Description = "Returns/sets a graphic for an special header button when it is down."

   Set SpecialButtonPictureDown = m_ButtonPicture(5)

End Property

Public Property Let SpecialButtonPictureDown(ByVal NewSpecialButtonPictureDown As IPictureDisp)

   Set SpecialButtonPictureDown = NewSpecialButtonPictureDown

End Property

Public Property Set SpecialButtonPictureDown(ByVal NewSpecialButtonPictureDown As IPictureDisp)

   Call SetButtonPicture(5, NewSpecialButtonPictureDown, "SpecialButtonPictureDown", True)

End Property

Public Property Get SpecialButtonPictureHover() As IPictureDisp
Attribute SpecialButtonPictureHover.VB_Description = "Returns/sets a graphic for an special header button when the mouse is hover the header."

   Set SpecialButtonPictureHover = m_ButtonPicture(4)

End Property

Public Property Let SpecialButtonPictureHover(ByVal NewSpecialButtonPictureHover As IPictureDisp)

   Set SpecialButtonPictureHover = NewSpecialButtonPictureHover

End Property

Public Property Set SpecialButtonPictureHover(ByVal NewSpecialButtonPictureHover As IPictureDisp)

   Call SetButtonPicture(4, NewSpecialButtonPictureHover, "SpecialButtonPictureHover", True)

End Property

Public Property Get SpecialButtonPictureUp() As IPictureDisp
Attribute SpecialButtonPictureUp.VB_Description = "Returns/sets a graphic for an special header button when it is up."

   Set SpecialButtonPictureUp = m_ButtonPicture(3)

End Property

Public Property Let SpecialButtonPictureUp(ByVal NewSpecialButtonPictureUp As IPictureDisp)

   Set SpecialButtonPictureUp = NewSpecialButtonPictureUp

End Property

Public Property Set SpecialButtonPictureUp(ByVal NewSpecialButtonPictureUp As IPictureDisp)

   Call SetButtonPicture(3, NewSpecialButtonPictureUp, "SpecialButtonPictureUp", True)

End Property

Public Property Get SpecialButtonUpColor() As OLE_COLOR
Attribute SpecialButtonUpColor.VB_Description = "Returns/sets the color for an special header button when it is up."

   SpecialButtonUpColor = m_SpecialButtonUpColor

End Property

Public Property Let SpecialButtonUpColor(ByVal NewSpecialButtonUpColor As OLE_COLOR)

   m_SpecialButtonUpColor = NewSpecialButtonUpColor
   PropertyChanged "SpecialButtonUpColor"
   
   Call Refresh

End Property

Public Property Get SpecialHeaderBackColor() As OLE_COLOR
Attribute SpecialHeaderBackColor.VB_Description = "Returns/sets the background color for an special header."

   SpecialHeaderBackColor = m_SpecialHeaderBackColor

End Property

Public Property Let SpecialHeaderBackColor(ByVal NewSpecialHeaderBackColor As OLE_COLOR)

   m_SpecialHeaderBackColor = NewSpecialHeaderBackColor
   PropertyChanged "SpecialHeaderBackColor"
   
   Call Refresh

End Property

Public Property Get SpecialHeaderForeColor() As OLE_COLOR
Attribute SpecialHeaderForeColor.VB_Description = "Returns/sets the foreground color for an special header."

   SpecialHeaderForeColor = m_SpecialHeaderForeColor

End Property

Public Property Let SpecialHeaderForeColor(ByVal NewSpecialHeaderForeColor As OLE_COLOR)

   m_SpecialHeaderForeColor = NewSpecialHeaderForeColor
   PropertyChanged "SpecialHeaderForeColor"
   
   Call Refresh

End Property

Public Property Get SpecialHeaderHoverColor() As OLE_COLOR
Attribute SpecialHeaderHoverColor.VB_Description = "Returns/sets the color for an special header when the mouse is hover the header."

   SpecialHeaderHoverColor = m_SpecialHeaderHoverColor

End Property

Public Property Let SpecialHeaderHoverColor(ByVal NewSpecialHeaderHoverColor As OLE_COLOR)

   m_SpecialHeaderHoverColor = NewSpecialHeaderHoverColor
   PropertyChanged "SpecialHeaderHoverColor"
   
   Call Refresh

End Property

Public Property Get SpecialItemBackColor() As OLE_COLOR
Attribute SpecialItemBackColor.VB_Description = "Returns/sets the background color for an special item group."

   SpecialItemBackColor = m_SpecialItemBackColor

End Property

Public Property Let SpecialItemBackColor(ByVal NewSpecialItemBackColor As OLE_COLOR)

   m_SpecialItemBackColor = NewSpecialItemBackColor
   PropertyChanged "SpecialItemBackColor"
   
   Call Refresh

End Property

Public Property Get SpecialItemBorderColor() As OLE_COLOR
Attribute SpecialItemBorderColor.VB_Description = "Returns/sets the border color for an special item group."

   SpecialItemBorderColor = m_SpecialItemBorderColor

End Property

Public Property Let SpecialItemBorderColor(ByVal NewSpecialItemBorderColor As OLE_COLOR)

   m_SpecialItemBorderColor = NewSpecialItemBorderColor
   PropertyChanged "SpecialItemBorderColor"
   
   Call Refresh

End Property

Public Property Get SpecialItemForeColor() As OLE_COLOR
Attribute SpecialItemForeColor.VB_Description = "Returns/sets the foreground color for an special item group."

   SpecialItemForeColor = m_SpecialItemForeColor

End Property

Public Property Let SpecialItemForeColor(ByVal NewSpecialItemForeColor As OLE_COLOR)

   m_SpecialItemForeColor = NewSpecialItemForeColor
   PropertyChanged "SpecialItemForeColor"
   
   Call Refresh

End Property

Public Property Get SpecialItemHoverColor() As OLE_COLOR
Attribute SpecialItemHoverColor.VB_Description = "Returns/sets the color for an special item group when the mouse is hover the item."

   SpecialItemHoverColor = m_SpecialItemHoverColor

End Property

Public Property Let SpecialItemHoverColor(ByVal NewSpecialItemHoverColor As OLE_COLOR)

   m_SpecialItemHoverColor = NewSpecialItemHoverColor
   PropertyChanged "SpecialItemHoverColor"
   
   Call Refresh

End Property

Public Property Get UseAlphaBlend() As Boolean
Attribute UseAlphaBlend.VB_Description = "Determins whether alphablend will be used by animations."

   UseAlphaBlend = m_UseAlphaBlend

End Property

Public Property Let UseAlphaBlend(ByVal NewUseAlphaBlend As Boolean)

   m_UseAlphaBlend = NewUseAlphaBlend
   PropertyChanged "UseAlphaBlend"

End Property

Public Property Get UseTheme() As ExploreBarThemeTypes
Attribute UseTheme.VB_Description = "Determins whether Windows themes will be used."

   UseTheme = m_UseTheme

End Property

Public Property Let UseTheme(ByVal NewUseTheme As ExploreBarThemeTypes)

   m_UseTheme = NewUseTheme
   PropertyChanged "UseTheme"
   
   Call Refresh

End Property

Public Property Get UseUserForeColors() As Boolean
Attribute UseUserForeColors.VB_Description = "Determins whether user colors will be used."

   UseUserForeColors = m_UseUserForeColors

End Property

Public Property Let UseUserForeColors(ByVal NewUseCustomTextColors As Boolean)

   m_UseUserForeColors = NewUseCustomTextColors
   PropertyChanged "UseUserForeColors"
   
   Call Refresh

End Property

Public Function AddDetailGroup(ByVal Title As String, Optional ByVal TitleBold As Boolean = True, Optional ByVal WindowState As WindowStates, Optional ByVal Icon As IPictureDisp, Optional ByVal DetailPicture As IPictureDisp, Optional ByVal DetailTitle As String, Optional ByVal DetailCaption As String, Optional ByVal ToolTipText As String, Optional ByVal Tag As String) As Integer

   If GroupDetailsSet Then
      AddDetailGroup = -1
      
   ElseIf GroupsCount = 20 Then
      AddDetailGroup = -20
      
   Else
      AddDetailGroup = AddGroup(Title, TitleBold, WindowState, Icon, , ToolTipText, Tag, , , True, DetailTitle, DetailCaption, DetailPicture)
   End If

End Function

Public Function AddItem(ByVal Group As Integer, ByVal Caption As String, Optional ByVal CaptionBold As Boolean, Optional ByVal Icon As IPictureDisp, Optional ByVal TextOnly As Boolean, Optional ByVal OpenFile As String, Optional ByVal Tag As String, Optional ByVal ToolTipText As String) As Integer

   AddItem = -1
   
   If Not CheckGroupExist(Group) Then Exit Function
   
   With Groups(Group)
      If .ContainerSet Then
         AddItem = -2
         Exit Function
         
      ElseIf .TypeGroup = GROUP_DETAILS Then
         AddItem = -3
         Exit Function
         
      ElseIf .ItemsCount = 20 Then
         AddItem = -20
         Exit Function
      End If
      
      ReDim Preserve .Items(.ItemsCount) As GroupItemType
      
      With .Items(.ItemsCount)
         .Caption = Caption
         .Bold = CaptionBold
         .TextOnly = TextOnly
         .OpenFile = OpenFile
         .Tag = Tag
         .ToolTipText = ToolTipText
         Set .Icon = Icon
      End With
      
      .Items(.ItemsCount).State = STATE_NORMAL
      AddItem = .ItemsCount
      .ItemsCount = .ItemsCount + 1
      
      Call Refresh
   End With

End Function

Public Function AddNormalGroup(ByVal Title As String, Optional ByVal TitleBold As Boolean = True, Optional ByVal WindowState As WindowStates, Optional ByVal Icon As IPictureDisp, Optional ByVal ItemsBackground As IPictureDisp, Optional ByVal ToolTipText As String, Optional ByVal Tag As String, Optional ByVal Container As PictureBox) As Integer

   If GroupsCount = 20 Then
      AddNormalGroup = -20
      
   Else
      AddNormalGroup = AddGroup(Title, TitleBold, WindowState, Icon, ItemsBackground, ToolTipText, Tag, Container)
   End If

End Function

Public Function AddSpecialGroup(ByVal Title As String, Optional ByVal TitleBold As Boolean = True, Optional ByVal WindowState As WindowStates, Optional ByVal Icon As IPictureDisp, Optional ByVal ItemsBackground As IPictureDisp, Optional ByVal ToolTipText As String, Optional ByVal Tag As String, Optional ByVal Container As PictureBox) As Integer

   If GroupSpecialsSet Then
      AddSpecialGroup = -1
      
   ElseIf GroupsCount = 20 Then
      AddSpecialGroup = 20
      
   Else
      AddSpecialGroup = AddGroup(Title, TitleBold, WindowState, Icon, ItemsBackground, ToolTipText, Tag, Container, True)
   End If

End Function

Public Function DeleteGroup(ByVal Group As Integer) As Integer

Dim intGroup       As Integer
Dim lngWindow      As Long
Dim lngWindowState As WindowStates

   DeleteGroup = -1
   
   If Not CheckGroupExist(Group) Then Exit Function
   
   With Groups(Group)
      If .TypeGroup = GROUP_SPECIAL Then GroupSpecialsSet = False
      If .TypeGroup = GROUP_DETAILS Then GroupDetailsSet = False
      
      If Not .Container Is Nothing Then
         Call Subclass_DelMsg(.Container.hWnd, WM_LBUTTONDOWN)
         Call Subclass_Stop(.Container.hWnd)
         
         Set .Container = Nothing
      End If
   End With
   
   For intGroup = Group To UBound(Groups) - 1
      lngWindowState = Groups(intGroup + 1).WindowState
      Groups(intGroup) = Groups(intGroup + 1)
      Groups(intGroup).WindowState = lngWindowState
      
      If lngWindowState = Collapsed Then picItems(intGroup).Visible = False
   Next 'intGroup
   
   If intGroup > 0 Then
      ReDim Preserve Groups(intGroup - 1) As GroupType
      
   Else
      Erase Groups
      GroupsExist = False
   End If
   
   GroupsCount = GroupsCount - 1
   DeleteGroup = GroupsCount
   lngWindow = picGroup.Item(intGroup).hWnd
   
   Call Subclass_DelMsg(lngWindow, WM_MOUSEWHEEL)
   Call Subclass_Stop(lngWindow)
   
   lngWindow = picItems.Item(intGroup).hWnd
   
   Call Subclass_DelMsg(lngWindow, WM_MOUSEWHEEL)
   Call Subclass_Stop(lngWindow)
   
   If GroupsCount Then
      Unload picGroup.Item(picGroup.Count - 1)
      Unload picItems.Item(picItems.Count - 1)
      
   Else
      picGroup.Item(0).Picture = Nothing
      picGroup.Item(0).Visible = False
      picItems.Item(0).Picture = Nothing
      picItems.Item(0).Visible = False
   End If
   
   GroupsExist = SafeArrayGetDim(Groups())
   
   Call Refresh

End Function

Public Function DeleteItem(ByVal Group As Integer, ByVal Item As Integer) As Integer

Dim intItem As Integer

   DeleteItem = -1
   
   If Not CheckGroupExist(Group) Then Exit Function
   
   With Groups(Group)
      If Not CheckItem(Group, Item) Or (.TypeGroup = GROUP_DETAILS) Or .ContainerSet Then Exit Function
      
      For intItem = Item To UBound(.Items) - 1
         .Items(intItem) = .Items(intItem + 1)
      Next 'intItem
      
      If intItem > 0 Then
         ReDim Preserve .Items(intItem - 1) As GroupItemType
         
      Else
         .WindowState = Collapsed
         picItems.Item(Group).Picture = Nothing
         picItems.Item(Group).Visible = False
         Erase .Items
      End If
      
      .ItemsCount = .ItemsCount - 1
      DeleteItem = .ItemsCount
      
      Call Refresh
   End With

End Function

Public Function FullTextShowed(ByVal Group As Integer, Optional ByVal Item As Integer = -1) As Boolean

   If Not CheckGroupExist(Group) Then Exit Function
   
   If Item > -1 Then
      If Not CheckItem(Group, Item) Or (Groups(Group).TypeGroup = GROUP_DETAILS) Then Exit Function
      
      FullTextShowed = Groups(Group).Items(Item).FullTextShowed
      
   Else
      FullTextShowed = Groups(Group).FullTextShowed
   End If

End Function

Public Function GetDetailCaption(ByVal Group As Integer) As String

   If Not CheckGroupExist(Group) Or (Groups(Group).TypeGroup <> GROUP_DETAILS) Then Exit Function
   
   GetDetailCaption = Groups(Group).DetailCaption

End Function

Public Function GetDetailPicture(ByVal Group As Integer) As IPictureDisp

   If Not CheckGroupExist(Group) Or (Groups(Group).TypeGroup <> GROUP_DETAILS) Then Exit Function
   
   Set GetDetailPicture = Groups(Group).DetailPicture

End Function

Public Function GetDetailTitle(ByVal Group As Integer) As String

   If Not CheckGroupExist(Group) Or (Groups(Group).TypeGroup <> GROUP_DETAILS) Then Exit Function
   
   GetDetailTitle = Groups(Group).DetailTitle

End Function

Public Function GetGroupContainer(ByVal Group As Integer) As PictureBox

   If Not CheckGroupExist(Group) Then Exit Function
   
   GetGroupContainer = Groups(Group).Container

End Function

Public Function GetGroupIcon(ByVal Group As Integer) As IPictureDisp

   If Not CheckGroupExist(Group) Then Exit Function
   
   Set GetGroupIcon = Groups(Group).Icon

End Function

Public Function GetGroupsCount() As Integer

   GetGroupsCount = GroupsCount

End Function

Public Function GetGroupState(ByVal Group As Integer) As Integer

   If Not CheckGroupExist(Group) Then Exit Function
   
   GetGroupState = Groups(Group).State

End Function

Public Function GetGroupTag(ByVal Group As Integer) As String

   If Not CheckGroupExist(Group) Then Exit Function
   
   GetGroupTag = Groups(Group).Tag

End Function

Public Function GetGroupTitle(ByVal Group As Integer) As String

   If Not CheckGroupExist(Group) Then Exit Function
   
   GetGroupTitle = Groups(Group).Title

End Function

Public Function GetGroupTitleBold(ByVal Group As Integer) As Boolean

   If Not CheckGroupExist(Group) Then Exit Function
   
   GetGroupTitleBold = Groups(Group).Bold

End Function

Public Function GetGroupToolTipText(ByVal Group As Integer) As String

   If Not CheckGroupExist(Group) Then Exit Function
   
   GetGroupToolTipText = Groups(Group).ToolTipText

End Function

Public Function GetGroupWindowState(ByVal Group As Integer) As WindowStates

   GetGroupWindowState = -1
   
   If Not CheckGroupExist(Group) Then Exit Function
   
   GetGroupWindowState = Groups(Group).WindowState

End Function

Public Function GetItemCaption(ByVal Group As Integer, ByVal Item As Integer, Optional ByVal StripTab As Boolean) As String

Dim strCaption As String

   If Not CheckGroupExist(Group) Or Not CheckItem(Group, Item) Or (Groups(Group).TypeGroup = GROUP_DETAILS) Then Exit Function
   
   strCaption = Groups(Group).Items(Item).Caption
   
   If StripTab And InStr(strCaption, vbTab) Then strCaption = Replace(strCaption, vbTab, " ", 1)
   
   GetItemCaption = strCaption

End Function

Public Function GetItemCaptionBold(ByVal Group As Integer, ByVal Item As Integer) As Boolean

   If Not CheckGroupExist(Group) Or Not CheckItem(Group, Item) Or (Groups(Group).TypeGroup = GROUP_DETAILS) Then Exit Function
   
   GetItemCaptionBold = Groups(Group).Items(Item).Bold

End Function

Public Function GetItemIcon(ByVal Group As Integer, ByVal Item As Integer) As IPictureDisp

   If Not CheckGroupExist(Group) Or Not CheckItem(Group, Item) Or (Groups(Group).TypeGroup = GROUP_DETAILS) Then Exit Function
   
   Set GetItemIcon = Groups(Group).Items(Item).Icon

End Function

Public Function GetItemOpenFile(ByVal Group As Integer, ByVal Item As Integer) As String

   If Not CheckGroupExist(Group) Or Not CheckItem(Group, Item) Or (Groups(Group).TypeGroup = GROUP_DETAILS) Then Exit Function
   
   GetItemOpenFile = Groups(Group).Items(Item).OpenFile

End Function

Public Function GetItemsCount(ByVal Group As Integer) As Integer

   GetItemsCount = -1
   
   If Not CheckGroupExist(Group) Or (Groups(Group).TypeGroup = GROUP_DETAILS) Then Exit Function
   
   GetItemsCount = Groups(Group).ItemsCount

End Function

Public Function GetItemsBackgroundPicture(ByVal Group As Integer) As IPictureDisp

   If Not CheckGroupExist(Group) Or (Groups(Group).TypeGroup = GROUP_DETAILS) Then Exit Function
   
   Set GetItemsBackgroundPicture = Groups(Group).ItemsBackgroundPicture

End Function

Public Function GetItemTag(ByVal Group As Integer, ByVal Item As Integer) As String

   If Not CheckGroupExist(Group) Or Not CheckItem(Group, Item) Or (Groups(Group).TypeGroup = GROUP_DETAILS) Then Exit Function
   
   GetItemTag = Groups(Group).Items(Item).Tag

End Function

Public Function GetItemTextOnly(ByVal Group As Integer, ByVal Item As Integer) As Boolean

   If Not CheckGroupExist(Group) Or Not CheckItem(Group, Item) Or (Groups(Group).TypeGroup = GROUP_DETAILS) Then Exit Function
   
   GetItemTextOnly = Groups(Group).Items(Item).TextOnly

End Function

Public Function GetItemToolTipText(ByVal Group As Integer, ByVal Item As Integer) As String

   If Not CheckGroupExist(Group) Or Not CheckItem(Group, Item) Or (Groups(Group).TypeGroup = GROUP_DETAILS) Then Exit Function
   
   GetItemToolTipText = Groups(Group).Items(Item).ToolTipText

End Function

Public Function GetThemeMap() As String

   GetThemeMap = ThemeMap

End Function

Public Function GetThemeName() As String

   GetThemeName = ThemeName

End Function

Public Function GetThemeColorMap() As String

   GetThemeColorMap = GetThemeColorMap

End Function

Public Function GetThemeColorName() As String

   GetThemeColorName = ThemeColorName

End Function

Public Function GetVersion() As String

   GetVersion = VERSION

End Function

Public Function hWnd() As Long

   hWnd = UserControl.hWnd

End Function

Public Function SetDetailCaption(ByVal Group As Integer, ByVal NewCaption As String) As Boolean

   SetDetailCaption = SetGroupProperties(IsDetailCaption, Group, NewCaption)

End Function

Public Function SetDetailPicture(ByVal Group As Integer, Optional ByVal NewPicture As IPictureDisp) As Boolean

   SetDetailPicture = SetGroupProperties(IsDetailPicture, Group, , , , NewPicture)

End Function

Public Function SetDetailTitle(ByVal Group As Integer, ByVal NewTitle As String) As Boolean

   SetDetailTitle = SetGroupProperties(IsDetailTitle, Group, NewTitle)

End Function

Public Function SetGroupContainer(ByVal Group As Integer, Optional ByVal NewContainer As PictureBox) As Boolean

Dim intGroup As Integer

   If Not Groups(Group).ContainerSet Then
      For intGroup = 0 To GroupsCount - 1
         If Groups(intGroup).ContainerSet Or (Groups(intGroup).TypeGroup = GROUP_DETAILS) Then Exit Function
      Next 'intGroup
   End If
   
   SetGroupContainer = SetGroupProperties(IsGroupContainer, Group, , , , , NewContainer)

End Function

Public Function SetGroupIcon(ByVal Group As Integer, Optional ByVal NewIcon As IPictureDisp) As Boolean

   SetGroupIcon = SetGroupProperties(IsGroupIcon, Group, , , , NewIcon)

End Function

Public Function SetGroupTag(ByVal Group As Integer, Optional ByVal NewTag As String) As Boolean

   SetGroupTag = SetGroupProperties(IsGroupTag, Group, NewTag)

End Function

Public Function SetGroupTitle(ByVal Group As Integer, ByVal NewTitle As String) As Boolean

   SetGroupTitle = SetGroupProperties(IsGroupTitle, Group, NewTitle)

End Function

Public Function SetGroupTitleBold(ByVal Group As Integer, ByVal NewTitleBold As Boolean) As Boolean

   SetGroupTitleBold = SetGroupProperties(IsGroupTitleBold, Group, , NewTitleBold)

End Function

Public Function SetGroupToolTipText(ByVal Group As Integer, Optional ByVal Item As Integer, Optional ByVal NewToolTipText As String) As Boolean

   SetGroupToolTipText = SetGroupProperties(IsToolTipText, Group, NewToolTipText)

End Function

Public Function SetGroupWindowState(ByVal Group As Integer, Optional ByVal NewWindowState As WindowStates) As Boolean

   If (NewWindowState < Expanded) Or (NewWindowState > Fixed) Then Exit Function
   
   If (Groups(Group).TypeGroup = GROUP_DETAILS) And Not m_DetailGroupButton Then
      Exit Function
      
   ElseIf (Groups(Group).TypeGroup <> GROUP_DETAILS) And (Groups(Group).WindowState = Fixed) Then
      Exit Function
   End If
   
   Do While tmrAnimation.Enabled
      DoEvents
   Loop
   
   SetGroupWindowState = SetGroupProperties(IsWindowState, Group, , NewWindowState)

End Function

Public Function SetItemCaption(ByVal Group As Integer, ByVal Item As Integer, ByVal NewCaption As String) As Boolean

   SetItemCaption = SetItemProperties(IsItemCaption, Group, Item, NewCaption)

End Function

Public Function SetItemCaptionBold(ByVal Group As Integer, ByVal Item As Integer, Optional ByVal NewCaptionBold As Boolean) As Boolean

   SetItemCaptionBold = SetItemProperties(IsItemCaptionBold, Group, Item, , NewCaptionBold)

End Function

Public Function SetItemIcon(ByVal Group As Integer, ByVal Item As Integer, Optional ByVal NewIcon As IPictureDisp) As Boolean

   SetItemIcon = SetItemProperties(IsItemIcon, Group, Item, , , NewIcon)

End Function

Public Function SetItemOpenFile(ByVal Group As Integer, ByVal Item As Integer, Optional ByVal NewOpenFile As String) As Boolean

   SetItemOpenFile = SetItemProperties(IsItemOpenFile, Group, Item, NewOpenFile)

End Function

Public Function SetItemsBackgroundPicture(ByVal Group As Integer, Optional ByVal NewBackgroundPicture As IPictureDisp) As Boolean

   SetItemsBackgroundPicture = SetGroupProperties(IsGroupItemsBackgroundPicture, Group, , , NewBackgroundPicture)

End Function

Public Function SetItemTag(ByVal Group As Integer, ByVal Item As Integer, Optional ByVal NewTag As String) As Boolean

   SetItemTag = SetItemProperties(IsItemTag, Group, Item, NewTag)

End Function

Public Function SetItemTextOnly(ByVal Group As Integer, ByVal Item As Integer, Optional ByVal NewTextOnly As Boolean) As Boolean

   SetItemTextOnly = SetItemProperties(IsItemTextOnly, Group, Item, , NewTextOnly)

End Function

Public Function SetItemToolTipText(ByVal Group As Integer, Optional ByVal Item As Integer, Optional ByVal NewToolTipText As String) As Boolean

   SetItemToolTipText = SetItemProperties(IsToolTipText, Group, Item, NewToolTipText)

End Function

Public Sub DeleteAllGroupItems(ByVal Group As Integer)

   If Not CheckGroupExist(Group) Then Exit Sub
   
   With Groups(Group)
      If .TypeGroup <> GROUP_DETAILS Then
         .ItemsCount = 0
         
         ReDim .Items(.ItemsCount) As GroupItemType
         
         Call Refresh
      End If
   End With

End Sub

Public Sub DeleteAllGroups()

   Do While GroupsCount
      DeleteGroup GroupsCount - 1
      DoEvents
      
      Call Refresh
   Loop
   
   StateSpecialGroup = False
   StateNormalGroup = False
   GroupsExist = False
   Erase Groups
   
   Call Refresh

End Sub

Public Sub Refresh()

Dim blnBold As Boolean
Dim intSize As Integer
Dim lngMaxY As Long
Dim sngSize As Single
Dim strText As String

   If m_Locked Or tmrAnimation.Enabled Or Not Extender.Parent.Visible Then Exit Sub
   
   Call DrawBackground
   
   With tsbVertical
      If Ambient.UserMode And GroupsExist Then
         lngMaxY = DrawGroups
         intSize = bdrBorder.Visible
         
         If (lngMaxY - 2) > ScaleHeight Then
            .Top = 0 - intSize
            .Height = ScaleHeight + intSize * 2
            
            If Not .Visible Then
               .Left = ScaleWidth - .Width + intSize
               .Visible = True
               .Max = GetMaxY(DrawGroups)
               
            ElseIf (ScaleHeight <> .Height) Or (GetMaxY(lngMaxY) <> .Max) Then
               .Max = GetMaxY(lngMaxY)
            End If
            
         ElseIf .Visible Then
            .Visible = False
            .Left = ScaleWidth
            DrawGroups
         End If
         
      Else
         .Visible = False
         blnBold = FontBold
         sngSize = FontSize
         FontBold = True
         FontSize = 12
         ForeColor = &HFFFFFF
         strText = THEME_NAME & " - v" & VERSION
         CurrentX = (ScaleWidth - TextWidth(strText)) / 2
         CurrentY = (ScaleHeight - TextHeight(strText)) / 2 - TextHeight(strText)
         Print strText
         FontBold = blnBold
         FontSize = sngSize
      End If
   End With

End Sub

Private Function AddGroup(ByVal Title As String, ByVal TitleBold As Boolean, ByVal WindowState As WindowStates, ByVal Icon As IPictureDisp, Optional ByVal ItemsBackground As IPictureDisp, Optional ByVal ToolTipText As String, Optional ByVal Tag As String, Optional ByVal Container As PictureBox, Optional ByVal SpecialGroup As Boolean, Optional ByVal DetailGroup As Boolean, Optional ByVal DetailTitle As String, Optional ByVal DetailCaption As String, Optional ByVal DetailPicture As IPictureDisp) As Integer

Dim intGroup     As Integer
Dim lngTypeGroup As Long

   If SpecialGroup Then
      lngTypeGroup = GROUP_SPECIAL
      GroupSpecialsSet = True
      
      If WindowState = Expanded Then StateSpecialGroup = True
      
   Else
      lngTypeGroup = GROUP_NORMAL
      
      If (WindowState = Expanded) And m_OpenOneGroupOnly Then
         If StateSpecialGroup Or StateNormalGroup Then
            WindowState = Collapsed
            
         Else
            StateNormalGroup = True
         End If
      End If
      
      If (DetailGroup And (WindowState <> Collapsed)) And m_DetailGroupButton Then
         If m_OpenOneGroupOnly Then
            WindowState = Collapsed
            
         Else
            StateNormalGroup = True
         End If
         
         If (WindowState = Fixed) And m_DetailGroupButton Then WindowState = Expanded
      End If
   End If
   
   ReDim Preserve Groups(GroupsCount) As GroupType
   
   GroupsExist = SafeArrayGetDim(Groups())
   intGroup = GroupsCount
   
   If GroupsCount Then
      Load picGroup.Item(GroupsCount)
      Load picItems.Item(GroupsCount)
   End If
   
   With picGroup.Item(GroupsCount)
      .Font = UserControl.Font
      Subclass_Initialize .hWnd
      
      Call Subclass_AddMsg(.hWnd, WM_MOUSEWHEEL)
   End With
   
   With picItems.Item(GroupsCount)
      .Font = UserControl.Font
      Subclass_Initialize .hWnd
      
      Call Subclass_AddMsg(.hWnd, WM_MOUSEWHEEL)
   End With
   
   If lngTypeGroup = GROUP_SPECIAL Then
      With Groups(0)
         If .TypeGroup <> GROUP_SPECIAL Then
            For intGroup = GroupsCount To 1 Step -1
               Groups(intGroup) = Groups(intGroup - 1)
               picItems.Item(intGroup).TabStop = True
               
               With Groups(intGroup - 1)
                  If Not .Container Is Nothing Then
                     SetParent .Container.hWnd, picItems.Item(intGroup).hWnd
                     picItems.Item(intGroup).TabStop = False
                  End If
                  
                  .DetailCaption = ""
                  .DetailTitle = ""
                  .ItemsCount = 0
                  .ContainerSet = False
                  Set .Icon = Nothing
                  Set .Container = Nothing
                  Set .DetailPicture = Nothing
               End With
            Next 'intGroup
         End If
      End With
      
   ElseIf lngTypeGroup = GROUP_NORMAL Then
      If (Groups(GroupsCount).TypeGroup <> GROUP_DETAILS) Then
         For intGroup = 0 To GroupsCount
            With Groups(intGroup)
               If (.TypeGroup = GROUP_DETAILS) And (.TypeGroup <> GROUP_SPECIAL) Then
                  Groups(GroupsCount) = Groups(intGroup)
                  picItems.Item(GroupsCount).TabStop = False
                  Exit For
               End If
            End With
         Next 'intGroup
         
         If intGroup > GroupsCount Then intGroup = GroupsCount
      End If
   End If
   
   With Groups(intGroup)
      If DetailGroup Then
         .DetailTitle = DetailTitle
         .DetailCaption = DetailCaption
         lngTypeGroup = GROUP_DETAILS
         GroupDetailsSet = True
         
         If Not DetailPicture Is Nothing Then Set .DetailPicture = DetailPicture
      End If
      
      ReDim .Items(0) As GroupItemType
      
      .State = STATE_NORMAL
      .WindowState = WindowState
      .TypeGroup = lngTypeGroup
      .Title = Title
      .Bold = TitleBold
      .ToolTipText = ToolTipText
      .Tag = Tag
      .ItemsCount = 0
      .ContainerSet = SetContainer(intGroup, Container)
      
      If .ContainerSet Or (.TypeGroup = GROUP_DETAILS) Then
         picItems.Item(intGroup).TabStop = False
         
      Else
         picItems.Item(intGroup).TabStop = True
      End If
      
      Set .Icon = Icon
      Set .ItemsBackgroundPicture = ItemsBackground
      AddGroup = intGroup
      GroupsCount = GroupsCount + 1
   End With
   
   If m_OpenOneGroupOnly Then
      Groups(0).WindowState = Expanded
      
      For intGroup = 1 To GroupsCount - 1
         With Groups(intGroup)
            If .TypeGroup = GROUP_NORMAL Then
               .WindowState = Collapsed
               
            ElseIf (.TypeGroup = GROUP_DETAILS) And m_DetailGroupButton Then
               .WindowState = Collapsed
            End If
         End With
      Next 'intGroup
   End If
   
   Call Refresh

End Function

Private Function ChangeColor(ByVal Color As OLE_COLOR, ByVal Lighter As Boolean) As OLE_COLOR

Dim intBlue  As Integer
Dim intGreen As Integer
Dim intRed   As Integer

   intRed = Val("&H" & Right(CStr(Hex(Color)), 2))
   intGreen = Val("&H" & Mid(CStr(Hex(Color)), 3, 2))
   intBlue = Val("&H" & Left(CStr(Hex(Color)), 2))
   
   If Len(CStr(Hex(Color))) = 4 Then intGreen = Val("&H" & Left(CStr(Hex(Color)), 2))
   If Len(CStr(Hex(Color))) = 2 Then intGreen = 0
   If Len(CStr(Hex(Color))) < 5 Then intBlue = 0
   
   If Lighter Then
      intRed = intRed + 33
      intGreen = intGreen + 49
      intBlue = intBlue + 57
      
      If intRed > 255 Then intRed = 255
      If intGreen > 255 Then intGreen = 255
      If intBlue > 255 Then intBlue = 255
      
   Else
      intRed = intRed - 33
      intGreen = intGreen - 49
      intBlue = intBlue - 57
      
      If intRed < 0 Then intRed = 0
      If intGreen < 0 Then intGreen = 0
      If intBlue < 0 Then intBlue = 0
   End If
   
   ChangeColor = RGB(intRed, intGreen, intBlue)

End Function

Private Function CheckGroupExist(ByVal Group As Integer) As Boolean

   If Not GroupsExist Then Exit Function
   If (Group >= 0) And (Group <= UBound(Groups)) Then CheckGroupExist = True

End Function

Private Function CheckItem(ByVal Group As Integer, ByVal Item As Integer) As Boolean

   If (Item >= 0) And (Item <= UBound(Groups(Group).Items)) Then CheckItem = True

End Function

Private Function CheckMouseInContainer(ByVal hWnd As Long, ByVal Window As Long) As Boolean

Dim blnFound As Boolean
Dim intGroup As Integer

   For intGroup = 0 To GroupsCount - 1
      If Groups(intGroup).ContainerSet Then
         If Groups(intGroup).Container.hWnd = Window Then
            blnFound = True
            Exit For
         End If
         
         If Extender.Container.hWnd = hWnd Then
            blnFound = True
            Exit For
         End If
      End If
   Next 'intGroup
   
   If blnFound And Not KeyPressed Then
      MouseHoverGroup = intGroup
      MouseHoverControl = True
      RaiseEvent MouseHover(intGroup, -1, Groups(intGroup).FullTextShowed)
      CheckMouseInContainer = True
   End If

End Function

Private Function CheckMouseInControl() As Long

Dim intGroup  As Integer
Dim lngWindow As Long
Dim ptaMouse  As PointAPI

   If GroupsExist Then
      GetCursorPos ptaMouse
      lngWindow = WindowFromPoint(ptaMouse.X, ptaMouse.Y)
      CheckMouseInControl = lngWindow
      
      For intGroup = 0 To UBound(Groups)
         If (lngWindow = picGroup.Item(intGroup).hWnd) Or (lngWindow = picItems.Item(intGroup).hWnd) Or (lngWindow = tsbVertical.hWnd) Or (lngWindow = UserControl.hWnd) Then
            CheckMouseInControl = UserControl.hWnd
            Exit For
         End If
      Next 'intGroup
   End If

End Function

Private Function CountButtonPictures() As Boolean

Dim intCount As Integer

   For intCount = 0 To 5
      If m_ButtonPicture(intCount) Is Nothing Then Exit Function
   Next 'intCount
   
   CountButtonPictures = True

End Function

Private Function DrawDetails(ByVal Group As Integer) As Long

Dim blnNotEmpty    As Boolean
Dim lngBaseY       As Long
Dim lngHeight      As Long
Dim lngPictureSize As Long
Dim lngY           As Long
Dim rctText        As Rect
Dim strCaption     As String
Dim strTitle       As String

   With picItems.Item(Group)
      lngBaseY = ITEM_SPACE
      lngPictureSize = 100 And (Not Groups(Group).DetailPicture Is Nothing)
      strTitle = Groups(Group).DetailTitle
      strCaption = MakeMessage(Groups(Group).DetailCaption, tsbVertical.Left - 40)
      .Cls
      .FontUnderline = False
      
      If m_UseUserForeColors Then
         .ForeColor = m_DetailsForeColor
         
      Else
         .ForeColor = ForeColorNormalItems
      End If
      
      If Len(strCaption) Or Len(strTitle) Or lngPictureSize Then
         If lngPictureSize Then lngHeight = 12
         If Len(strTitle) Then lngHeight = lngHeight + TextHeight(strTitle) + 9
         
         If Len(strCaption) Then
            lngHeight = lngHeight + TextHeight(strCaption) + 9
            
            If lngPictureSize Then lngHeight = lngHeight + ITEM_SPACE - 2
            If Len(strTitle) Then lngHeight = lngHeight - 7
         End If
         
         blnNotEmpty = True
         .Top = picGroup.Item(Group).Top + picGroup.Item(Group).Height
         .Left = GROUP_SPACE
         .Width = tsbVertical.Left - GROUP_SPACE * 2
         .Height = lngHeight + ITEM_SPACE + lngPictureSize
         rctText.Left = ITEM_SPACE
         rctText.Right = .ScaleWidth - rctText.Left
      End If
      
      Call DrawThemePart(.hDC, EBP_NORMALGROUPBACKGROUND, STATE_NORMAL, ObjectRect(picItems.Item(Group)))
      
      If m_UseTheme = Enhanced Then picItems.Item(Group).Line (0, -1)-(.ScaleWidth - 1, .ScaleHeight - 1), BorderColorItems, B
      If Not HasTheme Then Call DrawWindow(Group, m_NormalItemBorderColor, m_NormalItemBackColor, m_GradientNormalItemBackColor)
      If Not Groups(Group).ItemsBackgroundPicture Is Nothing Then .PaintPicture Groups(Group).ItemsBackgroundPicture, 1, 1, .ScaleWidth - 1, .ScaleHeight - 1, , , , , vbSrcAnd
      
      If lngPictureSize Then
         On Error Resume Next
         .PaintPicture Groups(Group).DetailPicture, (.ScaleWidth - lngPictureSize) / 2, lngBaseY, lngPictureSize, lngPictureSize, , , , , vbSrcAnd
         lngBaseY = lngBaseY + lngPictureSize + ITEM_SPACE * 2
         On Error GoTo 0
      End If
      
      If Len(strTitle) Then
         rctText.Top = lngBaseY
         rctText.Bottom = lngBaseY + TextHeight(strTitle)
         .FontBold = True
         DrawText .hDC, strTitle, -1, rctText, DT_LEFT Or DT_WORD_ELLIPSIS
         lngBaseY = lngBaseY + ITEM_SPACE * 2
      End If
      
      If Len(strCaption) Then
         rctText.Top = lngBaseY
         rctText.Bottom = .ScaleHeight
         .FontBold = False
         DrawText .hDC, strCaption, -1, rctText, DT_LEFT Or DT_WORD_ELLIPSIS
      End If
      
      If blnNotEmpty Then
         lngY = .Height
         
         Call SetAnimationExpand(Group)
         
      Else
         lngY = 0
         .Visible = False
      End If
   End With
   
   DrawDetails = lngY

End Function

Private Function DrawGroup(ByVal Group As Integer, Optional ByVal Y As Long) As Long

Dim intHeight As Integer
Dim intTemp   As Integer
Dim lngPartID As Long
Dim objGroup  As Object
Dim rctFrame  As Rect

   ReDim lngColor(2) As Long
   
   If ThemeName = THEME_LUNA Then
      intHeight = 24
      
   Else
      intHeight = 22
   End If
   
   If Y = 0 Then Y = picGroup.Item(Group).Top
   
   intHeight = intHeight + ((32 - intHeight) And (m_HeaderHeight = High))
   GroupIndex = Group
   Set objGroup = picGroup.Item(Group)
   
   If Groups(Group).TypeGroup = GROUP_SPECIAL Then
      lngColor(0) = m_GradientSpecialHeaderBackColor
      lngColor(1) = m_SpecialHeaderBackColor
      lngColor(2) = lngColor(1)
      lngPartID = EBP_SPECIALGROUPHEAD
      
   Else
      lngColor(0) = m_NormalHeaderBackColor
      lngColor(1) = m_GradientNormalHeaderBackColor
      lngColor(2) = lngColor(0)
      lngPartID = EBP_NORMALGROUPHEAD
   End If
   
   With objGroup
      .Top = Y
      .Left = GROUP_SPACE
      .Width = tsbVertical.Left - GROUP_SPACE * 2
      .Height = intHeight
      rctFrame.Right = .ScaleWidth
      rctFrame.Bottom = .ScaleHeight
      
      If DoRoundGroup Then Call RoundGroup(Group)
      
      Call DrawThemePart(.hDC, lngPartID, STATE_NORMAL, rctFrame)
   End With
   
   If Not HasTheme Then
      With rctFrame
         .Right = tsbVertical.Left - GROUP_SPACE
         .Bottom = .Top + intHeight
         objGroup.Line (.Left, .Top)-(.Right - 1, .Bottom - 1), lngColor(2), BF
         
         Call DrawGradient(objGroup.hDC, rctFrame, lngColor(0), lngColor(1), m_GradientStyle, Groups(Group).TypeGroup <> GROUP_SPECIAL)
      End With
   End If
   
   If m_UseTheme = Enhanced Then
      If Groups(Group).TypeGroup <> GROUP_SPECIAL Then
         intTemp = AlphaPercent
         AlphaPercent = 50
         
         Call DoAlphaBlend(picGroupMasker.hDC, picGroupMasker, objGroup)
         
         AlphaPercent = intTemp
      End If
   End If
   
   Call DrawGroupTitle(Group)
   Call DrawButton(Group)
   
   If (Group = PressedGroup) And (FocussedItem < 0) Then
      With rctFrame
         .Top = .Top + 2
         .Left = .Left + 2 + (32 And Not Groups(Group).Icon Is Nothing)
         .Right = .Right - 2
         .Bottom = .Bottom - 2
      End With
      
      objGroup.ForeColor = vbBlack
      DrawFocusRect objGroup.hDC, rctFrame
   End If
   
   If m_UseTheme = Enhanced Then
      With rctFrame
         lngColor(0) = picGroup(Group).Point(.Right - 2, .Top)
         picGroup(Group).Line (0, .Top)-(0, .Bottom), lngColor(0)
         picGroup(Group).Line (.Right - 1, .Top)-(.Right - 1, .Bottom), lngColor(0)
      End With
   End If
   
   Call DrawGroupIcon(Group)
   
   objGroup.Visible = True
   DrawGroup = Y + intHeight
   Set objGroup = Nothing
   Erase lngColor

End Function

Private Function DrawGroups() As Long

Dim intGroup As Integer
Dim lngY     As Long
Dim lngValue As Long

   lngValue = ((tsbVertical.Value * 10) And tsbVertical.Visible)
   lngY = GROUP_SPACE - lngValue
   
   DoRoundGroup = True
   
   For intGroup = 0 To UBound(Groups)
      With Groups(intGroup)
         lngY = DrawGroup(intGroup, lngY)
         
         If .WindowState <> Collapsed Then
            If .TypeGroup = GROUP_DETAILS Then
               lngY = lngY + DrawDetails(intGroup)
               
            Else
               lngY = lngY + DrawItems(intGroup)
            End If
         End If
      End With
      
      lngY = lngY + GROUP_SPACE
   Next 'intGroup
   
   DoRoundGroup = False
   DrawGroups = lngY + lngValue

End Function

Private Function DrawItems(ByVal Group As Integer) As Long

Const DT_RIGHT                   As Long = &H2
Const EBP_SPECIALGROUPBACKGROUND As Long = 9

Dim ctlControl                   As Control
Dim intItem                      As Integer
Dim intScaleMode                 As Integer
Dim lngBackColor                 As Long
Dim lngBorderColor               As Long
Dim lngContainerColor            As Long
Dim lngForeColor                 As Long
Dim lngGradientColor             As Long
Dim lngParentWindow              As Long
Dim lngPartID                    As Long
Dim lngX                         As Long
Dim lngY                         As Long
Dim objItems                     As Object
Dim rctText                      As Rect
Dim strCaption()                 As String

   lngY = ITEM_SPACE
   Set objItems = picItems.Item(Group)
   
   ReDim lngColor(1) As Long
   
   With objItems
      If Groups(Group).TypeGroup = GROUP_SPECIAL Then
         lngColor(0) = m_SpecialItemForeColor
         lngColor(1) = m_SpecialItemHoverColor
         lngBackColor = m_SpecialItemBackColor
         lngBorderColor = m_SpecialItemBorderColor
         lngGradientColor = m_GradientSpecialItemBackColor
         lngPartID = EBP_SPECIALGROUPBACKGROUND
         
      Else
         lngColor(0) = m_NormalItemForeColor
         lngColor(1) = m_NormalItemHoverColor
         lngBackColor = m_NormalItemBackColor
         lngBorderColor = m_NormalItemBorderColor
         lngGradientColor = m_GradientNormalItemBackColor
         lngPartID = EBP_NORMALGROUPBACKGROUND
      End If
      
      If Groups(Group).ContainerSet Then
         .FontBold = False
         .Top = picGroup.Item(Group).Top + picGroup.Item(Group).Height
         .Left = GROUP_SPACE
         .Width = tsbVertical.Left - GROUP_SPACE * 2
         .Picture = Nothing
         .Cls
         
         If Not Groups(Group).Container Is Nothing Then
            intScaleMode = Groups(Group).ContainerScaleMode
            On Local Error Resume Next
            
            With Groups(Group).Container
               .Top = ScaleY(1, vbPixels, intScaleMode)
               .Left = ScaleX(1, vbPixels, intScaleMode)
               .Width = ScaleX((tsbVertical.Left - GROUP_SPACE * 2) - 2, vbPixels, intScaleMode)
               objItems.Height = ScaleY(.Height, intScaleMode, vbPixels) + 2
               .Picture = Nothing
               .Cls
            End With
         End If
         
         With Groups(Group).Container
            If m_UseUserForeColors Then
               lngForeColor = lngColor(0)
               
            Else
               lngForeColor = ForeColorNormalItems
            End If
            
            .Visible = False
            .BackColor = lngBackColor
            lngParentWindow = .Parent.hWnd
            
            Call DrawThemePart(.hDC, lngPartID, STATE_NORMAL, ObjectRect(picItems.Item(Group), True))
            
            lngContainerColor = .Point(2, 2)
            
            For Each ctlControl In Extender.Container.Controls
               With ctlControl
                  If .Container.hWnd <> lngParentWindow Then
                     Set .Font = UserControl.Font
                     .FontBold = UserControl.FontBold
                     .ForeColor = lngForeColor
                     
                     If TypeOf ctlControl Is CheckBox Then .BackColor = lngContainerColor
                     If TypeOf ctlControl Is CommandButton Then .BackColor = lngContainerColor
                     If TypeOf ctlControl Is Frame Then .BackColor = lngContainerColor
                     If TypeOf ctlControl Is Label Then .BackStyle = vbTransparent
                     If TypeOf ctlControl Is OptionButton Then .BackColor = lngContainerColor
                     If (TypeOf ctlControl Is PictureBox) And (.Name <> Groups(Group).Container.Name) Then .BackColor = lngContainerColor
                  End If
               End With
            Next 'ctlControl
            
            On Local Error GoTo 0
            Set ctlControl = Nothing
            .Visible = True
         End With
         
      Else
         .FontBold = False
         .Top = picGroup.Item(Group).Top + picGroup.Item(Group).Height
         .Left = GROUP_SPACE
         .Width = tsbVertical.Left - GROUP_SPACE * 2
         .Height = Groups(Group).ItemsCount * (ITEM_SPACE * 2) + GROUP_SPACE
         .Picture = Nothing
         .Cls
         
         Call DrawThemePart(.hDC, lngPartID, STATE_NORMAL, ObjectRect(picItems.Item(Group)))
      End If
      
      If m_UseTheme = Enhanced Then objItems.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), BorderColorItems, B
      If Not HasTheme Then Call DrawWindow(Group, lngBorderColor, lngBackColor, lngGradientColor)
   End With
   
   If Not Groups(Group).ItemsBackgroundPicture Is Nothing Then
      With picAnimation
         .AutoSize = True
         .Picture = Groups(Group).ItemsBackgroundPicture
         TransparentBlt objItems.hDC, 1, 1, objItems.ScaleWidth - 2, objItems.ScaleHeight - 2, .hDC, 0, 0, .ScaleWidth, .ScaleHeight, .Point(0, 0)
         .AutoSize = False
         .Picture = Nothing
      End With
   End If
   
   If Groups(Group).ItemsCount > 0 Then
      For intItem = 0 To UBound(Groups(Group).Items)
         If Not Groups(Group).ContainerSet Then
            With Groups(Group).Items(intItem)
               If Not .Icon Is Nothing Then
                  lngX = ITEM_SPACE * 2 + 16
                  objItems.PaintPicture .Icon, ITEM_SPACE, lngY, 16, 16
                  
               Else
                  lngX = ITEM_SPACE
               End If
               
               If .State = STATE_HOT Then
                  If m_UseUserForeColors Then
                     objItems.ForeColor = lngColor(1)
                     
                  Else
                     objItems.ForeColor = ForeColorHoverItems
                  End If
                  
                  objItems.FontUnderline = True And Not .TextOnly
                  
               ' State = STATE_NORMAL or State = STATE_PRESSED
               Else
                  If m_UseUserForeColors Then
                     objItems.ForeColor = lngColor(0)
                     
                  Else
                     objItems.ForeColor = ForeColorNormalItems
                  End If
                  
                  objItems.FontUnderline = Len(.OpenFile)
               End If
               
               objItems.FontBold = .Bold
               rctText.Top = lngY + (16 - TextHeight(.Caption)) \ 2
               rctText.Bottom = rctText.Top + TextHeight(.Caption)
               .FullTextShowed = True
               
               If InStr(.Caption, vbTab) Then
                  strCaption = Split(.Caption, vbTab, 2)
                  
                  With rctText
                     .Left = lngX
                     FontBold = objItems.FontBold
                     .Right = (objItems.ScaleWidth - ITEM_SPACE) / 2 - 2
                     DrawText objItems.hDC, strCaption(0), -1, rctText, DT_LEFT Or DT_WORD_ELLIPSIS
                     
                     If Len(strCaption(0)) And (objItems.TextWidth(strCaption(0)) > rctText.Right - rctText.Left) Then Groups(Group).Items(intItem).FullTextShowed = False
                     
                     .Left = .Right + 2
                     .Right = objItems.ScaleWidth - ITEM_SPACE
                     DrawText objItems.hDC, strCaption(1), -1, rctText, DT_RIGHT Or DT_WORD_ELLIPSIS
                     lngX = .Right
                     
                     If Len(strCaption(1)) And (objItems.TextWidth(strCaption(1)) > .Right - .Left) Then Groups(Group).Items(intItem).FullTextShowed = False
                  End With
                  
               Else
                  rctText.Left = lngX
                  FontBold = objItems.FontBold
                  rctText.Right = objItems.ScaleWidth - ITEM_SPACE
                  DrawText objItems.hDC, .Caption, -1, rctText, DT_LEFT Or DT_WORD_ELLIPSIS
                  lngX = lngX + TextWidth(.Caption)
                  
                  If Len(.Caption) And (objItems.TextWidth(.Caption) > rctText.Right - rctText.Left) Then .FullTextShowed = False
               End If
               
               If (Group = PressedGroup) And (intItem = FocussedItem) Then
                  lngContainerColor = objItems.ForeColor
                  rctText.Left = ITEM_SPACE - 1
                  rctText.Right = lngX + 1
                  objItems.ForeColor = vbBlack
                  DrawFocusRect objItems.hDC, rctText
                  objItems.ForeColor = lngContainerColor
               End If
               
               FontBold = False
               .Rect.Top = lngY
               .Rect.Left = ITEM_SPACE
               .Rect.Right = lngX
               .Rect.Bottom = .Rect.Top + TextHeight(.Caption)
               lngY = lngY + ITEM_SPACE * 2
            End With
         End If
      Next 'intItem
      
      Call SetAnimationExpand(Group)
      
      lngY = objItems.Height
      
   Else
      lngY = 0
   End If
   
   DrawItems = lngY
   Set objItems = Nothing
   Erase lngColor, strCaption

End Function

Private Function GetColor(ByVal IsColor As Integer) As Integer

   GetColor = Val("&H" & Hex((IsColor / &HFF&) * &HFFFF&))

End Function

Private Function GetMaxY(ByVal Y As Long)

   GetMaxY = (Y - ScaleHeight + (2 And HasTheme)) \ 10

End Function

Private Function GetThemeTextColor(ByVal hTheme As Long, ByVal PartID As Long, ByVal State As Long) As Long

Const TMT_TEXTCOLOR As Integer = 3803

Dim lngColor As Long

   GetThemeColor hTheme, PartID, State, TMT_TEXTCOLOR, lngColor
   GetThemeTextColor = lngColor

End Function

Private Function IsFunctionSupported(ByVal sFunction As String, ByVal sModule As String) As Boolean

Dim lngModule As Long

   lngModule = GetModuleHandle(sModule)
   
   If lngModule = 0 Then lngModule = LoadLibrary(sModule)
   
   If lngModule Then
      IsFunctionSupported = GetProcAddress(lngModule, sFunction)
      FreeLibrary lngModule
   End If

End Function

Private Function MakeMessage(ByVal Message As String, ByVal Width As Long, Optional ByRef LinesCount As Integer) As String

Dim intLines    As Integer
Dim intPointer  As Integer
Dim lngCount    As Long
Dim strBuffer() As String
Dim strLine     As String
Dim strText     As String
Dim strWord     As String

   strBuffer = Split(Message)
   
   For lngCount = 0 To UBound(strBuffer)
      strWord = strBuffer(lngCount)
      
      If InStr(strWord, vbCrLf) Then
         intPointer = InStr(strWord, vbCrLf)
         strBuffer(lngCount) = Mid(strWord, intPointer)
         strWord = Left(strWord, intPointer - 1)
      End If
      
      If TextWidth(strText & " " & strWord) >= Width Then
         strLine = strLine & strText & vbCrLf
         intLines = intLines + 1
         strText = strWord
         
      Else
         strText = LTrim(strText & " " & strWord)
      End If
      
      If intPointer Then
         strLine = strLine & strText
         intPointer = 0
         
         Do
            strBuffer(lngCount) = Mid(strBuffer(lngCount), 3)
            strLine = strLine & vbCrLf
            intLines = intLines + 1
            
            If Left(strBuffer(lngCount), 2) <> vbCrLf Then
               lngCount = lngCount - 1
               strText = ""
               Exit Do
            End If
         Loop
      End If
   Next 'lngCount
   
   LinesCount = intLines + 1
   MakeMessage = strLine & strText

End Function

Private Function ObjectRect(ByRef Box As Object, Optional ByVal HasContainer As Boolean) As Rect

   With ObjectRect
      .Left = 1 + HasContainer
      .Right = Box.ScaleWidth - 1 + HasContainer
      .Bottom = Box.ScaleHeight - 1 + HasContainer
   End With

End Function

Private Function SetContainer(ByVal Group As Integer, ByVal Container As PictureBox) As Boolean

Dim ctlControl      As Control
Dim lngParentWindow As Long

   With Groups(Group)
      If Container Is Nothing Then
         If .Container Is Nothing Then Exit Function
         
         Call Subclass_DelMsg(.Container.hWnd, WM_LBUTTONDOWN)
         Call Subclass_Stop(.Container.hWnd)
         
         .ItemsCount = 0
         SetParent .Container.hWnd, Container.Parent.hWnd
         Set .Container = Nothing
         picItems.Item(Group) = Nothing
         picItems.Item(Group).Visible = False
         
         ReDim .Items(0) As GroupItemType
         
         Call Refresh
         
      Else
         ReDim .Items(0) As GroupItemType
         
         .ItemsCount = 1
         Set .Container = Container
         .ContainerScaleMode = .Container.Parent.ScaleMode
         .Items(0).State = STATE_NORMAL
         .Items(0).Tag = .Container.Tag
         .Items(0).ToolTipText = .Container.ToolTipText
         .Container.ScaleMode = vbPixels
         .Container.AutoRedraw = True
         .Container.TabStop = False
         lngParentWindow = .Container.Parent.hWnd
         SetParent .Container.hWnd, picItems.Item(Group).hWnd
         SetContainer = True
         Subclass_Initialize .Container.hWnd
         
         Call Subclass_AddMsg(.Container.hWnd, WM_LBUTTONDOWN)
      End If
      
      Set Container = Nothing
   End With

End Function

Private Function SetGroupProperties(ByVal GroupPropertie As GroupItemProperties, ByVal Group As Integer, Optional ByVal NewText As String, Optional ByVal NewBold As Boolean, Optional ByVal NewWindowState As WindowStates, Optional ByVal NewPicture As IPictureDisp, Optional ByVal NewContainer As PictureBox) As Boolean

Dim blnRefresh As Boolean

   If Not CheckGroupExist(Group) Then Exit Function
   
   On Local Error GoTo ExitFunction
   blnRefresh = True
   
   With Groups(Group)
      Select Case GroupPropertie
         Case IsGroupIcon
            Set .Icon = NewPicture
            
         Case IsGroupContainer
            SetContainer Group, Nothing
            .ContainerSet = SetContainer(Group, NewContainer)
            
         Case IsGroupTag
            blnRefresh = False
            .Tag = NewText
            
         Case IsGroupTitle
            .Title = NewText
            
         Case IsGroupTitleBold
            .Bold = NewBold
            
         Case IsGroupItemsBackgroundPicture
            Set .ItemsBackgroundPicture = NewPicture
            
         Case IsToolTipText
            blnRefresh = False
            .ToolTipText = NewText
            
         Case IsWindowState
            GroupAnimated = Group
            
            If NewWindowState = Collapsed Then
               .State = STATE_HOT
               
               Call SetAnimationCollapse(Group)
               
            Else
               .State = STATE_NORMAL
            End If
            
            .WindowState = NewWindowState
            
            If .WindowState = Expanded Then Call CheckOpenGroups(Group)
            
         Case Else
            If .TypeGroup <> GROUP_DETAILS Then GoTo ExitFunction
            
            Select Case GroupPropertie
               Case IsDetailCaption
                  .DetailCaption = NewText
                  
               Case IsDetailPicture
                  Set .DetailPicture = NewPicture
                  
               Case IsDetailTitle
                  .DetailTitle = NewText
            End Select
      End Select
   End With
   
   SetGroupProperties = True
   
   If blnRefresh Then Call Refresh
   
ExitFunction:
   On Local Error GoTo 0

End Function

Private Function SetItemProperties(ByVal ItemPropertie As GroupItemProperties, ByVal Group As Integer, ByVal Item As Integer, Optional ByVal NewText As String, Optional ByVal NewValue As Boolean, Optional ByVal NewPicture As IPictureDisp) As Boolean

Dim blnRefresh As Boolean

   If Not CheckGroupExist(Group) Then Exit Function
   
   On Local Error GoTo ExitFunction
   blnRefresh = True
   
   With Groups(Group).Items(Item)
      Select Case ItemPropertie
         Case IsItemCaption
            .Caption = NewText
            
         Case IsItemCaptionBold
            .Bold = NewValue
            
         Case IsItemIcon
            Set .Icon = NewPicture
            
         Case IsItemOpenFile
            .OpenFile = NewText
            
         Case IsItemTag
            blnRefresh = False
            .Tag = NewText
            
         Case IsItemTextOnly
            blnRefresh = False
            .TextOnly = NewValue
            
         Case IsToolTipText
            blnRefresh = False
            .ToolTipText = NewText
      End Select
   End With
   
   SetItemProperties = True
   
   If blnRefresh Then Call Refresh
   
ExitFunction:
   On Local Error GoTo 0

End Function

Private Function StripNull(ByVal Text As String) As String

   StripNull = Left(Text, StrLen(StrPtr(Text)))

End Function

Private Function TranslateColor(ByVal Colors As OLE_COLOR, Optional ByVal Palette As Long) As Long

   If OleTranslateColor(Colors, Palette, TranslateColor) Then TranslateColor = -1

End Function

Private Function TranslateRGB(ByVal ColorVal As Long, ByVal Part As Long) As Long

Dim strHex As String

   strHex = Trim(Hex(ColorVal))
   TranslateRGB = Val("&H" + UCase(Mid(Right("000000", 6 - Len(strHex)) & strHex, 5 - Part * 2, 2)))

End Function

Private Sub CheckFocus()

   If MouseHoverControl Then Exit Sub
   
   If Not MouseHoverScrollBar And (FocussedGroup > -1) Then
      Call ResetFocussedGroup
      
      FocussedGroup = -1
      FocussedItem = -1
   End If

End Sub

Private Sub CheckOpenGroups(ByVal Group As Integer)

Dim intGroup As Integer

   If (Groups(Group).WindowState = Fixed) Or Not m_OpenOneGroupOnly Then Exit Sub
   
   For intGroup = 0 To UBound(Groups)
      If intGroup <> Group Then
         If Groups(intGroup).WindowState = Expanded Then
            Groups(intGroup).State = STATE_NORMAL
            Groups(intGroup).WindowState = Collapsed
            GroupAnimated = intGroup
            
            Call SetAnimationCollapse(intGroup)
            
            Do While tmrAnimation.Enabled
               DoEvents
            Loop
         End If
      End If
   Next 'intGroup
   
   GroupAnimated = Group

End Sub

Private Sub DoAlphaBlend(ByVal hDC As Long, ByVal Source As PictureBox, ByVal Destination As PictureBox)

Const AC_SRC_OVER    As Long = &H0

Dim bftAlpha         As BlendFunctionType
Dim lngBlendFunction As Long

   With bftAlpha
      .BlendOp = AC_SRC_OVER
      .BlendFlags = 0
      .SourceConstantAlpha = AlphaPercent
      .AlphaFormat = AC_SRC_OVER
   End With
   
   Call CopyMemory(lngBlendFunction, bftAlpha, 4)
   
   With Destination
      AlphaBlend .hDC, 0, 0, .ScaleWidth, .ScaleHeight, hDC, Source.Left, Source.Top, Source.ScaleWidth, Source.ScaleHeight, lngBlendFunction
   End With

End Sub

Private Sub DrawArrow(ByVal Group As Integer, ByVal X As Long, ByVal Y As Long, ByVal IsWindowState As WindowStates, ByVal Color As Long)

Dim objGroup As Object

   Set objGroup = picGroup.Item(Group)
   
   If IsWindowState = Collapsed Then
      objGroup.Line (X + 1, Y + 11)-(X + 4, Y + 14), Color
      objGroup.Line (X + 4, Y + 12)-(X + 6, Y + 10), Color
      objGroup.Line (X, Y + 11)-(X + 4, Y + 15), Color
      objGroup.Line (X + 4, Y + 13)-(X + 7, Y + 10), Color
      objGroup.Line (X + 1, Y + 15)-(X + 4, Y + 18), Color
      objGroup.Line (X + 4, Y + 16)-(X + 6, Y + 14), Color
      objGroup.Line (X, Y + 15)-(X + 4, Y + 19), Color
      objGroup.Line (X + 4, Y + 17)-(X + 7, Y + 14), Color
      
   Else
      objGroup.Line (X + 1, Y + 13)-(X + 4, Y + 10), Color
      objGroup.Line (X + 4, Y + 12)-(X + 6, Y + 14), Color
      objGroup.Line (X, Y + 13)-(X + 4, Y + 9), Color
      objGroup.Line (X + 4, Y + 11)-(X + 7, Y + 14), Color
      objGroup.Line (X + 1, Y + 17)-(X + 4, Y + 14), Color
      objGroup.Line (X + 4, Y + 16)-(X + 6, Y + 18), Color
      objGroup.Line (X, Y + 17)-(X + 4, Y + 13), Color
      objGroup.Line (X + 4, Y + 15)-(X + 7, Y + 18), Color
   End If
   
   Set objGroup = Nothing

End Sub

Private Sub DrawBackground()

Dim lngBackColor        As Long
Dim lngGradientColor    As Long
Dim rctFrame            As Rect

   With rctFrame
      .Right = ScaleWidth
      .Bottom = ScaleHeight
   End With
   
   Cls
   
   Call DrawThemePart(hDC, EBP_HEADERBACKGROUND, STATE_NORMAL, rctFrame)
   
   If Not HasTheme Then
      Call DrawGradient(hDC, rctFrame, m_GradientBackColor, m_BackColor, m_GradientStyle)
      
   ElseIf m_UseTheme = Enhanced Then
      If ThemeName = THEME_EMBEDDED Then
         lngBackColor = &HAB6100
         lngGradientColor = &H754400
         
      ElseIf ThemeName = THEME_MEDIACENTRE Then
         lngBackColor = &HE8A062
         lngGradientColor = &HBE7253
         
      ElseIf ThemeName = THEME_ROYALE Then
         If ThemeColorName = "Metallic" Then
            lngBackColor = &HA6A399
            lngGradientColor = &H746A61
            
         Else
            lngBackColor = &HE8A062
            lngGradientColor = &HBE7253
         End If
         
      ElseIf ThemeName = THEME_ZUNE Then
         lngBackColor = &H808080
         lngGradientColor = &H454545
      End If
      
      If lngBackColor Then
         Call DrawGradient(hDC, rctFrame, lngGradientColor, lngBackColor, BottomTop)
      End If
   End If
   
   Picture = Image

End Sub

Private Sub DrawButton(ByVal Group As Integer)

Const EBP_NORMALGROUPCOLLAPSE  As Integer = 6
Const EBP_NORMALGROUPEXPAND    As Integer = 7
Const EBP_SPECIALGROUPCOLLAPSE As Integer = 10
Const EBP_SPECIALGROUPEXPAND   As Integer = 11

Dim intHeight                  As Integer
Dim lngCollapse                As Long
Dim lngColor(2)                As Long
Dim lngWindowState             As Long
Dim objGroup                   As Object
Dim rctButton                  As Rect

   If (Groups(Group).TypeGroup = GROUP_DETAILS) And Not m_DetailGroupButton Then Exit Sub
   
   Set objGroup = picGroup.Item(Group)
   intHeight = (4 And (m_HeaderHeight = High))
   objGroup.FillStyle = vbFSSolid
   
   With rctButton
      .Top = (picGroup.Item(Group).ScaleHeight - 24) / 2 - m_HeaderHeight
      .Left = tsbVertical.Left - 54
      .Right = .Left + 25
      .Bottom = .Top + 25
   End With
   
   With Groups(Group)
      If .TypeGroup = GROUP_SPECIAL Then
         lngCollapse = EBP_SPECIALGROUPCOLLAPSE
         lngWindowState = EBP_SPECIALGROUPEXPAND
         lngColor(0) = m_SpecialButtonUpColor
         lngColor(1) = m_SpecialButtonBackColor
         lngColor(2) = m_SpecialArrowUpColor
         
      Else
         lngCollapse = EBP_NORMALGROUPCOLLAPSE
         lngWindowState = EBP_NORMALGROUPEXPAND
         lngColor(0) = m_NormalButtonUpColor
         lngColor(1) = m_NormalButtonBackColor
         lngColor(2) = m_NormalArrowUpColor
      End If
      
      If (.WindowState <> Fixed) Or m_DetailGroupButton Then
         If .WindowState <> Collapsed Then lngWindowState = lngCollapse
         
         If (ThemeName = THEME_LUNA) Or (m_UseTheme = Windows) Then
            Call DrawThemePart(objGroup.hDC, lngWindowState, .State, rctButton)
            
         Else
            If (lngWindowState = EBP_NORMALGROUPEXPAND) Or (lngWindowState = EBP_SPECIALGROUPEXPAND) Then
               lngWindowState = 0
               
            Else
               lngWindowState = 1
            End If
            
            Call DrawButtonsThemed(Group, lngWindowState, .State - 1, intHeight, rctButton)
         End If
         
         If Not HasTheme Then
            If .State > STATE_NORMAL Then
               If .State = STATE_PRESSED Then FillStyle = vbFSSolid
               
               If .TypeGroup = GROUP_SPECIAL Then
                  If .State = STATE_PRESSED Then
                     lngColor(0) = m_SpecialButtonHoverColor
                     lngColor(1) = m_SpecialButtonDownColor
                     lngColor(2) = m_SpecialArrowDownColor
                     
                  Else
                     lngColor(0) = m_SpecialButtonHoverColor
                     lngColor(1) = m_SpecialButtonBackColor
                     lngColor(2) = m_SpecialArrowHoverColor
                  End If
                  
               Else
                  If .State = STATE_PRESSED Then
                     lngColor(0) = m_NormalButtonHoverColor
                     lngColor(1) = m_NormalButtonDownColor
                     lngColor(2) = m_NormalArrowDownColor
                     
                  Else
                     lngColor(0) = m_NormalButtonHoverColor
                     lngColor(2) = m_NormalArrowHoverColor
                  End If
               End If
            End If
            
            If .State Then
               If UseButtonPictures Then
                  objGroup.PaintPicture picButtonMasker, tsbVertical.Left - 53, 7 + intHeight, 19, 19, , , , , vbSrcAnd
                  objGroup.PaintPicture m_ButtonPicture(.State - 1 + (3 And (.TypeGroup = GROUP_SPECIAL))), tsbVertical.Left - 43, 7 + intHeight, 19, 19, , , , , vbSrcPaint
                  
               Else
                  objGroup.FillColor = lngColor(1)
                  objGroup.Circle (tsbVertical.Left - 43, intHeight + 10), 8, lngColor(0)
               End If
            End If
            
            objGroup.FillStyle = vbFSTransparent
            
            Call DrawArrow(Group, tsbVertical.Left - 46, intHeight - 4, .WindowState, lngColor(2))
         End If
         
      ElseIf (ThemeName <> THEME_LUNA) Or (m_UseTheme = Enhanced) Then
         Call DrawButtonsThemed(Group, 0, 0, 0, rctButton)
      End If
   End With
   
   Set objGroup = Nothing

End Sub

Private Sub DrawButtonsThemed(ByVal Group As Integer, ByVal WindowState As Long, ByVal State As Long, ByVal Height As Integer, ByRef ButtonRect As Rect)

Dim intX As Integer
Dim intY As Integer

   intX = WindowState * 7
   intY = State * 9
   TransparentBlt picGroup(Group).hDC, ButtonRect.Left + 8, ButtonRect.Top + 6 - m_HeaderHeight, 7, 10 + Height, picButtons.hDC, intX, intY, 7, 9, picButtons.Point(0, 0)

End Sub

Private Sub DrawGradient(ByVal hDC As Long, ByRef picRect As Rect, ByVal GradientColor As Long, ByVal BaseColor As Long, ByVal Style As GradientStyles, Optional ByVal IsNormalHeader As Boolean)

Dim blnSwap        As Boolean
Dim lngDirection   As Long
Dim lngRGB         As Long
Dim rctGradient    As GradientRect
Dim tvxGradient(1) As TriVertex

   If (Style = RightLeft) Or (Style = BottomTop) Then
      GradientColor = GradientColor Xor BaseColor
      BaseColor = GradientColor Xor BaseColor
      GradientColor = GradientColor Xor BaseColor
      blnSwap = True
   End If
   
   lngDirection = 1 And (Style < LeftRight)
   lngRGB = TranslateColor(GradientColor)
   
   With tvxGradient(0)
      If IsNormalHeader Then
         If blnSwap Then
            .X = picRect.Left
            picRect.Right = picRect.Left + 160
            
         Else
            .X = picRect.Right - 160
         End If
         
      Else
         .X = picRect.Left
      End If
      
      .Y = picRect.Top
      .Red = GetColor(TranslateRGB(lngRGB, IsRed))
      .Green = GetColor(TranslateRGB(lngRGB, IsGreen))
      .Blue = GetColor(TranslateRGB(lngRGB, IsBlue))
   End With
   
   lngRGB = TranslateColor(BaseColor)
   
   With tvxGradient(1)
      .X = picRect.Right
      .Y = picRect.Bottom
      .Red = GetColor(TranslateRGB(lngRGB, IsRed))
      .Green = GetColor(TranslateRGB(lngRGB, IsGreen))
      .Blue = GetColor(TranslateRGB(lngRGB, IsBlue))
   End With
   
   rctGradient.UpperLeft = 1
   rctGradient.LowerRight = 0
   GradientFill hDC, tvxGradient(0), 2, rctGradient, 1, lngDirection
   Erase tvxGradient

End Sub

Private Sub DrawGroupIcon(ByVal Group As Integer)

   If Groups(Group).Icon Is Nothing Then Exit Sub
   
   With picGroup.Item(Group)
      Cls
      .PaintPicture Groups(Group).Icon, -1, .ScaleHeight - 32, 32, 32
      PaintPicture Groups(Group).Icon, .Left - 1, .Top + .ScaleHeight - 32, 32, 32
   End With

End Sub

Private Sub DrawGroupTitle(ByVal Group As Integer)

Dim rctText  As Rect
Dim strTitle As String

   With picGroup.Item(Group)
      If (Groups(Group).State = STATE_NORMAL) Or (PressedGroup > -1) Then
         If Groups(Group).TypeGroup = GROUP_SPECIAL Then
            If m_UseUserForeColors Then
               .ForeColor = m_SpecialHeaderForeColor
               
            Else
               .ForeColor = ForeColorNormalGroupSpecial
            End If
            
         Else
            If m_UseUserForeColors Then
               .ForeColor = m_NormalHeaderForeColor
               
            Else
               .ForeColor = ForeColorNormalGroups
            End If
            
            If (Groups(Group).TypeGroup = GROUP_DETAILS) And Not m_DetailGroupButton Then
               If m_UseUserForeColors Then
                  .ForeColor = m_SpecialHeaderHoverColor
                  
               Else
                  .ForeColor = ForeColorHoverGroupSpecial
               End If
            End If
         End If
         
      ElseIf Groups(Group).TypeGroup = GROUP_SPECIAL Then
         If m_UseUserForeColors Then
            .ForeColor = m_SpecialHeaderHoverColor
            
         Else
            .ForeColor = ForeColorHoverGroupSpecial
         End If
         
      ElseIf m_UseUserForeColors Then
         .ForeColor = m_NormalHeaderHoverColor
         
      Else
         .ForeColor = ForeColorHoverGroups
      End If
      
      .FontBold = Groups(Group).Bold
      .FontUnderline = False
      strTitle = Groups(Group).Title
      rctText.Top = (.ScaleHeight - TextHeight(strTitle)) \ 2 - 1
      rctText.Left = 40 - (30 And Groups(Group).Icon Is Nothing)
      rctText.Right = tsbVertical.Left - 32 - (18 And (Groups(Group).WindowState <> Fixed))
      rctText.Bottom = .ScaleHeight - rctText.Top
      Groups(Group).FullTextShowed = True
      
      If Len(strTitle) And (.TextWidth(strTitle) > rctText.Right - rctText.Left) Then Groups(Group).FullTextShowed = False
      
      DrawText .hDC, strTitle, -1, rctText, DT_LEFT Or DT_WORD_ELLIPSIS
   End With

End Sub

Private Sub DrawThemePart(ByVal hDC As Long, ByVal PartID As Long, ByVal State As Long, ByRef FrameRect As Rect)

Const DEF_NORMAL        As Integer = 1
Const TBP_BACKGROUNDTOP As Integer = 1
Const TDP_FLASHBUTTON   As Integer = 2

Dim lngTheme            As Long
Dim strThemeName        As String
Dim rctFrame            As Rect

   If m_UseTheme = User Then GoTo ErrorTheme
   
   rctFrame = FrameRect
   
   If ((PartID = EBP_SPECIALGROUPHEAD) Or (PartID = EBP_NORMALGROUPHEAD)) And (m_UseTheme = Enhanced) Then
      If ThemeName = THEME_ZUNE Then
         strThemeName = "TaskBand"
         PartID = TDP_FLASHBUTTON
         
         With rctFrame
            .Top = .Top - 4
            .Left = -3
            .Right = .Right + 3
            .Bottom = .Bottom + 4
         End With
         
      Else
         strThemeName = "TaskBar"
         PartID = TBP_BACKGROUNDTOP
         
         With rctFrame
            .Left = 1
            .Right = .Right - 1
            .Bottom = .Bottom + 1
         End With
      End If
      
      State = DEF_NORMAL
      
   Else
      strThemeName = THEME_NAME
   End If
   
   On Local Error GoTo ErrorTheme
   lngTheme = OpenThemeData(UserControl.hWnd, StrPtr(strThemeName))
   
   If lngTheme Then
      HasTheme = Not DrawThemeBackground(lngTheme, hDC, PartID, State, rctFrame, rctFrame)
   End If
   
   CloseThemeData lngTheme
   
   GoTo ExitFunction
   
ErrorTheme:
   HasTheme = False
   
ExitFunction:
   On Local Error GoTo 0
   CloseThemeData lngTheme

End Sub

Private Sub DrawWindow(ByVal Group As Integer, ByVal IsBorderColor As Long, ByVal IsBackColor As Long, ByVal IsGradientColor As Long)

Dim rctFrame As Rect

   ReDim lngColor(1) As Long
   
   With rctFrame
      .Top = 1
      .Left = 1
      .Right = picItems.Item(Group).ScaleWidth - 2
      .Bottom = picItems.Item(Group).ScaleHeight - 2
      picItems.Item(Group).Line (0, 0)-(.Right + 1, .Bottom + 1), IsBorderColor, B
   End With
   
   If Groups(Group).TypeGroup = GROUP_SPECIAL Then
      lngColor(1) = IsGradientColor
      lngColor(0) = IsBackColor
      
   Else
      lngColor(0) = IsGradientColor
      lngColor(1) = IsBackColor
   End If
   
   Call DrawGradient(picItems.Item(Group).hDC, rctFrame, lngColor(0), lngColor(1), m_GradientStyle)
   
   Erase lngColor

End Sub

Private Sub GetTextColors()

Dim bytThemeName(520)      As Byte
Dim bytThemeColorName(520) As Byte
Dim intPointer             As Integer
Dim lngTheme               As Long

   If m_UseTheme = User Then Exit Sub
   
   On Local Error GoTo ExitSub
   lngTheme = OpenThemeData(hWnd, StrPtr(THEME_NAME))
   
   If lngTheme Then
      GetCurrentThemeName VarPtr(bytThemeName(0)), 260, VarPtr(bytThemeColorName(0)), 260, 0, 0
      ThemeName = StripNull(CStr(bytThemeName))
      ThemeName = Left(ThemeName, InStrRev(ThemeName, "\") - 1)
      ThemeMap = ThemeName
      ThemeName = Mid(ThemeName, InStrRev(ThemeName, "\") + 1)
      ThemeColorName = StripNull(CStr(bytThemeColorName))
      ThemeColorMap = ThemeMap & "\Shell\" & ThemeColorName
      
      If (ThemeName = THEME_LUNA) Or (m_UseTheme <> Enhanced) Then
         ForeColorNormalGroupSpecial = GetThemeTextColor(lngTheme, EBP_SPECIALGROUPHEAD, STATE_NORMAL)
         ForeColorNormalGroups = GetThemeTextColor(lngTheme, EBP_NORMALGROUPHEAD, STATE_NORMAL)
         ForeColorHoverGroups = ChangeColor(ForeColorNormalGroups, True)
         ForeColorHoverGroupSpecial = ChangeColor(ForeColorNormalGroups, True)
         ForeColorNormalItems = ForeColorNormalGroups
         ForeColorHoverItems = ForeColorHoverGroups
         
         If m_UseTheme = Enhanced Then
            If ThemeColorName = "NormalColor" Then
               ForeColorNormalGroups = ChangeColor(ForeColorHoverGroups, True)
               ForeColorHoverGroups = ChangeColor(ForeColorNormalItems, False)
               
            Else
               ForeColorHoverGroups = GetThemeTextColor(lngTheme, EBP_NORMALGROUPHEAD, STATE_NORMAL)
               ForeColorNormalGroups = ChangeColor(ForeColorHoverGroups, True)
            End If
         End If
         
      Else
         ForeColorHoverGroupSpecial = GetThemeTextColor(lngTheme, EBP_SPECIALGROUPHEAD, STATE_NORMAL)
         ForeColorNormalGroupSpecial = ChangeColor(ForeColorHoverGroupSpecial, False)
         ForeColorHoverGroups = ForeColorHoverGroupSpecial
         ForeColorNormalGroups = ForeColorNormalGroupSpecial
         ForeColorNormalItems = GetThemeTextColor(lngTheme, EBP_NORMALGROUPBACKGROUND, STATE_NORMAL)
         ForeColorHoverItems = ChangeColor(ForeColorNormalItems, True)
         
         If ThemeName = THEME_ZUNE Or ThemeName = THEME_EMBEDDED Then
            BorderColorItems = vbWhite
            
         Else
            BorderColorItems = ForeColorNormalItems
         End If
      End If
   End If
   
ExitSub:
   On Local Error GoTo 0

End Sub

Private Sub MakeGroupVisible(ByVal Group As Integer)

   If (picGroup(Group).Top < GROUP_SPACE) Or ((picGroup(Group).Top + picGroup(Group).ScaleHeight + picItems(Group).ScaleHeight) > tsbVertical.Height) Then tsbVertical.Value = Group * 10
   If (Groups(Group).WindowState <> Collapsed) And (Groups(Group).TypeGroup <> GROUP_DETAILS) And picItems.Item(Group).Visible Then DrawItems Group

End Sub

Private Sub MouseMoveGroup(ByVal Group As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Optional ByVal MouseMove As Boolean)

Dim blnChanged  As Boolean
Dim blnNewMouse As Boolean

   With Groups(Group)
      If PtInRect(ObjectRect(picGroup(Group)), CLng(X), CLng(Y)) = 0 Then
         picGroup.Item(Group).ToolTipText = ""
         .State = STATE_NORMAL
         blnChanged = True
         
      Else
         Call SetToolTipText(picGroup.Item(Group), .ToolTipText)
         
         GroupIndex = Group
         
         If Button = vbDefault Then
            If .State <> STATE_HOT Then
               .State = STATE_HOT
               blnChanged = True
               
            Else
               blnChanged = True
               MouseMove = True
            End If
            
         ElseIf Button = vbLeftButton Then
            If .State <> STATE_PRESSED Then
               .State = STATE_PRESSED
               blnChanged = True
            End If
         End If
         
         blnNewMouse = True
      End If
      
      If .WindowState <> Fixed Then picGroup.Item(Group).MousePointer = (vbCustom And blnNewMouse)
      
      If blnChanged Then
         If MouseMove Then
            DrawGroup Group
            
         Else
            Call MakeGroupVisible(Group)
            Call Refresh
            
            If Groups(Group).WindowState = Collapsed Then
               RaiseEvent Collapse(Group)
               
            ElseIf Groups(Group).WindowState = Expanded Then
               RaiseEvent Expand(Group)
            End If
         End If
      End If
   End With

End Sub

Private Sub MouseMoveItem(ByVal Group As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Optional ByVal MouseMove As Boolean)

Dim blnChanged    As Boolean
Dim blnNewMouse   As Boolean
Dim intItem       As Integer
Dim strTooTipText As String

   GroupIndex = Group
   strTooTipText = picItems.Item(Group).ToolTipText
   
   With Groups(Group)
      Call ResetGroup
      
      If .TypeGroup <> GROUP_DETAILS Then
         If .WindowState <> Collapsed Then
            For intItem = 0 To UBound(.Items)
               With .Items(intItem)
                  If PtInRect(.Rect, CLng(X), CLng(Y)) Or (KeyPressed And (intItem = FocussedItem)) Then
                     If Button = vbLeftButton Then FocussedItem = intItem
                     
                     If .State <> STATE_HOT Then
                        If KeyPressed Then
                           .State = STATE_PRESSED
                           
                        Else
                           .State = STATE_HOT
                        End If
                        
                        blnChanged = True
                        
                        If intItem <> MouseHoverItem Then
                           RaiseEvent MouseHover(Group, intItem, .FullTextShowed)
                           MouseHoverItem = intItem
                        End If
                     End If
                     
                     If picItems.Item(Group).MousePointer <> vbCustom Then picItems.Item(Group).MousePointer = vbCustom
                     
                     If Len(.ToolTipText) Then
                        If .ToolTipText <> strTooTipText Then Call SetToolTipText(picItems.Item(Group), .ToolTipText)
                        
                     ElseIf Len(.OpenFile) Then
                        If .OpenFile <> strTooTipText Then Call SetToolTipText(picItems.Item(Group), .OpenFile)
                     End If
                     
                     blnNewMouse = True
                     
                  ElseIf .State <> STATE_NORMAL Then
                     .State = STATE_NORMAL
                     blnChanged = True
                     picItems.Item(Group).ToolTipText = ""
                     
                     If (Button <> vbLeftButton) And (intItem = MouseHoverItem) Then
                        MouseHoverItem = -1
                        RaiseEvent MouseOut(Group, intItem)
                     End If
                  End If
               End With
            Next 'intItem
         End If
      End If
   End With
   
   picItems.Item(Group).MousePointer = (vbCustom And blnNewMouse)
   
   If blnChanged Then
      If MouseMove Then
         DrawItems Group
         
      Else
         Call Refresh
      End If
   End If

End Sub

Private Sub MoveAll()

Dim intGroup As Integer
Dim lngY     As Long

   lngY = GROUP_SPACE - ((tsbVertical.Value * 10) And tsbVertical.Visible)
   
   For intGroup = 0 To UBound(Groups)
      With picGroup.Item(intGroup)
         .Top = lngY
         lngY = lngY + .Height
         
         If Groups(intGroup).TypeGroup = GROUP_SPECIAL Then
            Cls
            
            If Not Groups(intGroup).Icon Is Nothing Then PaintPicture Groups(intGroup).Icon, .Left - 1, .Top + .ScaleHeight - 32, 32, 32
         End If
         
         If picItems.Item(intGroup).Visible Then
            picItems.Item(intGroup).Top = .Top + .Height
            lngY = lngY + picItems.Item(intGroup).Height
         End If
      End With
      
      lngY = lngY + 15
   Next 'intGroup

End Sub

Private Sub MoveGroupItemWindow(ByVal Lines As Integer)

Dim intGroup As Integer
Dim lngTop   As Long

   AlphaPercent = AlphaPercent + BlendStep
   
   If AlphaPercent > 255 Then AlphaPercent = 255
   If AlphaPercent < 0 Then AlphaPercent = 0
   
   With picItems.Item(GroupAnimated)
      If Lines < 0 Then
         If m_UseAlphaBlend Then Call DoAlphaBlend(hDC, picItems.Item(GroupAnimated), picItems.Item(GroupAnimated))
         
         BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight - Lines, .hDC, 0, -Lines, vbSrcCopy
      End If
      
      With Groups(GroupAnimated)
         If Not .Container Is Nothing Then .Container.Top = .Container.Top + ScaleY(Lines, vbPixels, .ContainerScaleMode)
      End With
      
      .Height = .Height + Lines
      lngTop = .Top + .Height
      
      If Lines > 0 Then
         BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, picAnimation.hDC, 0, ItemHeight - .ScaleHeight, vbSrcCopy
         
         If m_UseAlphaBlend Then Call DoAlphaBlend(hDC, picItems.Item(GroupAnimated), picItems.Item(GroupAnimated))
      End If
      
      .Refresh
   End With
   
   For intGroup = GroupAnimated + 1 To UBound(Groups)
      With picGroup.Item(intGroup)
         .Top = lngTop + GROUP_SPACE
         picItems.Item(intGroup).Top = .Top + .Height
         
         Call DrawGroupIcon(intGroup)
         
         With picItems.Item(intGroup)
            If .Visible Then
               lngTop = .Top + .Height
               
            Else
               lngTop = .Top
            End If
         End With
         
         DoEvents
      End With
   Next 'intGroup

End Sub

Private Sub PlaySound(ByVal Index As Integer)

Const SND_ASYNC     As Long = &H1
Const SND_MEMORY    As Long = &H4
Const SND_NODEFAULT As Long = &H2

Dim strSoundBuffer  As String

   On Local Error Resume Next
   strSoundBuffer = StrConv(LoadResData(Index, "Sounds"), vbUnicode)
   SoundPlay strSoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
   On Local Error GoTo 0

End Sub

Private Sub ResetFocussedGroup()

   PressedGroup = -1
   DrawGroup FocussedGroup
   
   If picItems.Item(FocussedGroup).Visible And (Groups(FocussedGroup).TypeGroup <> GROUP_DETAILS) Then DrawItems FocussedGroup

End Sub

Private Sub ResetGroup()

   If Not GroupsExist Then Exit Sub
   
   Groups(GroupIndex).State = STATE_NORMAL
   DrawGroup GroupIndex

End Sub

Private Sub RoundGroup(ByVal Group As Integer)

Const RGN_OR     As Long = 2

Dim intCurve     As Integer
Dim lngRegion(1) As Long

   If UseTheme = Enhanced Then
      If ThemeName = THEME_MEDIACENTRE Or ((ThemeName = THEME_ROYALE) And (ThemeColorName = "NormalColor")) Then intCurve = 2
      
   Else
      intCurve = 4
   End If
   
   With picGroup(Group)
      lngRegion(0) = CreateRoundRectRgn(0, 0, .ScaleWidth + 1, .ScaleHeight + 1, intCurve, intCurve)
      lngRegion(1) = CreateRectRgn(0, intCurve, .ScaleWidth, .ScaleHeight)
      CombineRgn lngRegion(0), lngRegion(0), lngRegion(1), RGN_OR
      DeleteObject lngRegion(1)
      SetWindowRgn .hWnd, lngRegion(0), True
   End With
   
   DeleteObject lngRegion(0)
   Erase lngRegion
   UserControl.Refresh

End Sub

Private Sub SetAlignment()

   If Extender.Align < vbAlignLeft Then
      DoAlignment = True
      Extender.Align = vbAlignLeft
      
      If Width > Extender.Parent.ScaleWidth Then Width = Extender.Parent.ScaleWidth
   End If
   
   ' if alignment = vbAlignRight
   If DoAlignment And (Width > Extender.Parent.ScaleWidth) Then
      Width = Extender.Parent.ScaleWidth
      Extender.Align = vbAlignLeft
      DoAlignment = False
   End If

End Sub

Private Sub SetAnimationCollapse(ByVal Group As Integer)

   DrawGroup Group
   
   If m_Animation Then
      AlphaPercent = 0
      BlendStep = picGroup.Item(Group).ScaleHeight \ MoveLines
      
      If m_Animation = Slow Then
         BlendStep = BlendStep / 3.2
         
      ElseIf m_Animation = Medium Then
         BlendStep = BlendStep * 1.2
         
      ElseIf m_Animation = Fast Then
         BlendStep = BlendStep * 3.5
      End If
   End If
   
   Call ToggleAnimation(True)

End Sub

Private Sub SetAnimationExpand(ByVal Group As Integer)

   With picItems.Item(Group)
      If GroupAnimated = Group Then
         If m_Animation Then
            ItemHeight = .ScaleHeight
            
            Call MakeGroupVisible(Group)

            AlphaPercent = 255
            BlendStep = -255 \ (ItemHeight \ MoveLines)
            
            If m_Animation > Slow Then BlendStep = BlendStep / 1.5
            
            If Groups(Group).WindowState = Expanded Then
               Call ToggleAnimation(True)
               
               .Height = 0
               .Visible = True
               
               If Not Groups(Group).Container Is Nothing Then Groups(Group).Container.Top = -Groups(Group).Container.Height
            End If
            
         Else
            .Visible = True
         End If
         
      ElseIf FocussedGroup <> Group Then
         .Visible = True
      End If
   End With

End Sub

Private Sub SetButtonPicture(ByVal Index As Integer, ByVal NewPicture As IPictureDisp, ByVal PropertieSet As String, ByVal SpecialGroup As Boolean)

   Set m_ButtonPicture(Index) = NewPicture
   UseButtonPictures = CountButtonPictures
   PropertyChanged PropertieSet
   
   Call Refresh

End Sub

Private Sub SetToolTipText(ByVal Box As PictureBox, ByVal Text As String)

   Box.ToolTipText = Replace(Text, vbTab, " ")

End Sub

Private Sub ToggleAnimation(ByVal State As Boolean)

Dim intGroup As Integer
Dim intSpace As Integer

   If m_Animation Then
      tmrAnimation.Enabled = State
      
   Else
      picItems.Item(GroupIndex).Visible = False
      State = False
   End If
   
   If Not State Then
      With picItems.Item(GroupAnimated)
         intGroup = GroupAnimated + 1
         GroupAnimated = -1
         
         If m_OpenOneGroupOnly Then
            .Picture = picAnimation.Image
            
            If intGroup < GroupsCount Then
               If .Visible Then
                  intSpace = (picGroup.Item(intGroup).Top - (.Top + .Height)) - GROUP_SPACE
                  
               Else
                  intSpace = picGroup.Item(intGroup).Top - (picGroup.Item(intGroup - 1).Top + picGroup.Item(intGroup - 1).Height) - GROUP_SPACE
               End If
               
               For intGroup = intGroup To GroupsCount - 1
                  picGroup.Item(intGroup).Top = picGroup.Item(intGroup).Top - intSpace
                  picItems.Item(intGroup).Top = picItems.Item(intGroup).Top - intSpace
                  
                  If Not Groups(intGroup).Container Is Nothing Then Groups(intGroup).Container.Top = Groups(intGroup).Container.Top - ScaleY(intSpace, vbPixels, Groups(intGroup).ContainerScaleMode)
               Next 'intGroup
            End If
         End If
      End With
      
      If Not m_OpenOneGroupOnly Then Call Refresh
      
      Call MakeGroupVisible(FocussedGroup)
   End If
   
   KeyPressed = State

End Sub

Private Sub TrackMouseLeave(ByVal lhWnd As Long)

Const TME_LEAVE   As Long = &H2&

Dim tmeMouseTrack As TrackMouseEventStruct

   With tmeMouseTrack
      .cbSize = Len(tmeMouseTrack)
      .dwFlags = TME_LEAVE
      .hwndTrack = lhWnd
   End With
   
   If TrackUser32 Then
      TrackMouseEvent tmeMouseTrack
      
   Else
      TrackMouseEventComCtl tmeMouseTrack
   End If

End Sub

Private Sub picGroup_GotFocus(Index As Integer)

   If IsMouseDown Then
      IsMouseDown = False
      Exit Sub
   End If
   
   If FocussedGroup > -1 Then
      Call ResetFocussedGroup
      
      FocussedItem = -1
      FocussedGroup = Index
      PressedGroup = Index
      DrawGroup PressedGroup
      
      If picItems.Item(Index).Visible And (Groups(Index).TypeGroup <> GROUP_DETAILS) Then DrawItems Index
      
      Call MakeGroupVisible(Index)
   End If

End Sub

Private Sub picGroup_LostFocus(Index As Integer)

   Call CheckFocus

End Sub

Private Sub picGroup_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   IsMouseDown = True
   
   If tmrAnimation.Enabled Or (Button <> vbLeftButton) Then Exit Sub
   If (Groups(Index).TypeGroup <> GROUP_DETAILS) Or ((Groups(Index).TypeGroup = GROUP_DETAILS) And m_DetailGroupButton) Then RaiseEvent MouseDown(Index, Button, Shift, X, Y)
   If FocussedGroup > -1 Then Call ResetFocussedGroup
   
   FocussedGroup = Index
   FocussedItem = -1
   
   Call MouseMoveGroup(Index, Button, Shift, X, Y)

End Sub

Private Sub picGroup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If KeyPressed Or tmrAnimation.Enabled Then Exit Sub
   
   If Index <> MouseHoverGroup Then
      MouseHoverGroup = Index
      MouseHoverItem = -1
      ClickedItem = -1
      RaiseEvent MouseHover(Index, -1, Groups(Index).FullTextShowed)
      
   Else
      RaiseEvent MouseMove(Index, Button, Shift, X, Y)
   End If
   
   If Not FreezeMouseMove Then Call MouseMoveGroup(Index, Button, Shift, X, Y, True)

End Sub

Private Sub picGroup_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Const SOUND_GROUP_CLICKED As Integer = 1

   IsMouseDown = False
   
   With Groups(Index)
      If ((.ItemsCount = 0) And (.TypeGroup <> GROUP_DETAILS)) Or tmrAnimation.Enabled Or (Button <> vbLeftButton) Then Exit Sub
      If (.TypeGroup = GROUP_DETAILS) And (.DetailPicture Is Nothing) And (.DetailTitle = "") And (.DetailCaption = "") Then Exit Sub
   End With
   
   If PtInRect(ObjectRect(picGroup(Index)), CLng(X), CLng(Y)) Then
      If (Groups(Index).TypeGroup <> GROUP_DETAILS) Or ((Groups(Index).TypeGroup = GROUP_DETAILS) And m_DetailGroupButton) Then RaiseEvent MouseUp(Index, Button, Shift, X, Y)
      
   Else
      Groups(Index).State = STATE_NORMAL
      
      Call MouseMoveGroup(Index, 0, Shift, X, Y)
      
      Exit Sub
   End If
   
   With Groups(Index)
      If Button = vbLeftButton Then
         If (.WindowState <> Fixed) Or m_DetailGroupButton Then
            FreezeMouseMove = True
            GroupAnimated = Index
            
            If .WindowState = Expanded Then
               .State = STATE_HOT
               .WindowState = Collapsed
               
               Call SetAnimationCollapse(Index)
               
            Else
               .State = STATE_NORMAL
               .WindowState = Expanded
               
               Call CheckOpenGroups(Index)
            End If
         End If
         
         If (Groups(Index).TypeGroup <> GROUP_DETAILS) Or ((Groups(Index).TypeGroup = GROUP_DETAILS) And m_DetailGroupButton) Then
            If m_SoundGroupClicked And (Groups(Index).WindowState <> Fixed) Then Call PlaySound(SOUND_GROUP_CLICKED)
            
            RaiseEvent GroupClick(Index, .WindowState)
         End If
         
         Call MouseMoveGroup(Index, 0, Shift, X, Y)
         
         FreezeMouseMove = False
      End If
   End With

End Sub

Private Sub picItems_Click(Index As Integer)

Const SW_SHOWNORMAL As Long = &H1

Dim lngResult       As Long

   If (ClickedItem = -1) Or tmrAnimation.Enabled Then Exit Sub
   
   With Groups(Index).Items(ClickedItem)
      If Len(.OpenFile) Then
         lngResult = ShellExecute(hWnd, "Open", .OpenFile, "", "", SW_SHOWNORMAL)
         
         If lngResult < 32 Then
            RaiseEvent ErrorOpenFile(Index, ClickedItem, .OpenFile, lngResult)
            
         Else
            RaiseEvent ItemOpenFile(Index, ClickedItem, .OpenFile)
         End If
         
      Else
         RaiseEvent ItemClick(Index, ClickedItem)
      End If
   End With
   
   Groups(Index).Items(ClickedItem).State = STATE_NORMAL
   DrawItems Index
   ClickedItem = -1

End Sub

Private Sub picItems_GotFocus(Index As Integer)

   If IsMouseDown Then
      IsMouseDown = False
      Exit Sub
   End If

   If FocussedGroup > -1 Then
      Call ResetFocussedGroup
      
      If Index < FocussedGroup Then
         FocussedItem = Groups(Index).ItemsCount - 1
         
      Else
         FocussedItem = 0
      End If
      
      FocussedGroup = Index
      PressedGroup = Index
      
      If picItems.Item(Index).Visible And (Groups(Index).TypeGroup <> GROUP_DETAILS) Then DrawItems Index
      
      Call MakeGroupVisible(Index)
   End If

End Sub

Private Sub picItems_LostFocus(Index As Integer)

   Call CheckFocus

End Sub

Private Sub picItems_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   IsMouseDown = True
   
   If tmrAnimation.Enabled Or (Button <> vbLeftButton) Then Exit Sub
   If FocussedGroup > -1 Then Call ResetFocussedGroup
   
   FocussedGroup = Index
   RaiseEvent MouseDown(Index, Button, Shift, X, Y)
   
   Call MouseMoveItem(Index, Button, Shift, X, Y)

End Sub

Private Sub picItems_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If KeyPressed Or tmrAnimation.Enabled Then Exit Sub
   
   If Index <> MouseHoverGroup Then
      MouseHoverGroup = Index
      RaiseEvent MouseHover(Index, -1, Groups(Index).FullTextShowed)
   End If
   
   RaiseEvent MouseMove(Index, Button, Shift, X, Y)
   
   If Not FreezeMouseMove Then Call MouseMoveItem(Index, Button, Shift, X, Y, True)

End Sub

Private Sub picItems_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Const SOUND_ITEM_CLICKED As Integer = 2

Dim blnChanged           As Boolean
Dim intItem              As Integer

   IsMouseDown = False
   
   If tmrAnimation.Enabled Or (Button <> vbLeftButton) Then Exit Sub
   
   RaiseEvent MouseUp(Index, Button, Shift, X, Y)
   
   With Groups(Index)
      If .TypeGroup <> GROUP_DETAILS Then
         If .WindowState <> Collapsed Then
            If KeyPressed Then
               intItem = FocussedItem
               MouseHoverItem = FocussedItem
               blnChanged = True
               
            Else
               For intItem = 0 To UBound(.Items)
                  With .Items(intItem)
                     If PtInRect(.Rect, CLng(X), CLng(Y)) Then
                        If (Button = vbLeftButton) Then blnChanged = True
                        
                        Exit For
                     End If
                  End With
               Next 'intItem
            End If
            
            If blnChanged Then
               If Not .Items(intItem).TextOnly Then
                  If m_SoundItemClicked Then Call PlaySound(SOUND_ITEM_CLICKED)
                  
                  Call MouseMoveItem(Index, Button, 0, X, Y)
               End If
               
               ClickedItem = intItem
            End If
            
            If ClickedItem <> intItem Then
               If MouseHoverItem > -1 Then
                  RaiseEvent MouseOut(Index, MouseHoverItem)
                  MouseHoverItem = -1
                  
               Else
                  RaiseEvent GroupClick(Index, .WindowState)
               End If
               
            ElseIf .Items(ClickedItem).TextOnly Then
               ClickedItem = -1
            End If
         End If
      End If
   End With

End Sub

Private Sub tmrAnimation_Timer()

   If GroupAnimated = -1 Then
      tmrAnimation.Enabled = False
      Exit Sub
   End If
   
   With picItems.Item(GroupAnimated)
      If .Height < MoveLines Then
         With picAnimation
            .Cls
            .Width = picItems.Item(GroupAnimated).Width
            .Height = ItemHeight
            BitBlt .hDC, 0, 0, .ScaleWidth, ItemHeight, picItems.Item(GroupAnimated).hDC, 0, 0, vbSrcCopy
         End With
      End If
      
      If Groups(GroupAnimated).WindowState = Expanded Then
         If .Height + MoveLines < ItemHeight Then
            Call MoveGroupItemWindow(MoveLines)
            
         Else
            If Not Groups(GroupAnimated).Container Is Nothing Then Groups(GroupAnimated).Container.Top = ScaleY(1, vbPixels, Groups(GroupAnimated).ContainerScaleMode)
            
            .Height = ItemHeight
            
            Call ToggleAnimation(False)
            Call Refresh
         End If
      
      'WindowState = Collapsed
      ElseIf .Height - MoveLines > 0 Then
         Call MoveGroupItemWindow(-MoveLines)
         
      Else
         .Visible = False
         
         Call ToggleAnimation(False)
         
         If Not m_OpenOneGroupOnly Then Call Refresh
      End If
   End With

End Sub

Private Sub tsbVertical_Change()

   Call MoveAll

End Sub

Private Sub tsbVertical_LostFocus()

   Call CheckFocus

End Sub

Private Sub tsbVertical_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   tsbVertical.SetFocus
   
   If MouseHoverControl And Not MouseHoverScrollBar Then
      MouseHoverScrollBar = True
      RaiseEvent MouseHover(-2, -1, False)
   End If

End Sub

Private Sub tsbVertical_MouseWheel(ScrollLines As Integer)

   tsbVertical.Value = tsbVertical.Value + ScrollLines

End Sub

Private Sub tsbVertical_Scroll()

   Call MoveAll

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

   Call Refresh

End Sub

Private Sub UserControl_Initialize()

   ClickedItem = -1
   FocussedGroup = -1
   FocussedItem = -1
   GroupsExist = SafeArrayGetDim(Groups())
   GroupAnimated = -1
   MouseHoverGroup = -1
   MouseHoverItem = -1
   m_Animation = Medium
   m_BackColor = &HE6AA8C
   m_DetailsForeColor = &HC65D21
   m_GradientBackColor = &HCC6633
   m_GradientNormalHeaderBackColor = &HF0D2C5
   m_GradientNormalItemBackColor = &HF7DFD6
   m_GradientSpecialHeaderBackColor = &HB24801
   m_GradientSpecialItemBackColor = &HF7DFD6
   m_NormalArrowDownColor = &HC65D21
   m_NormalArrowHoverColor = &HC65D21
   m_NormalArrowUpColor = &HFF8E42
   m_NormalButtonBackColor = &HFFFFFF
   m_NormalButtonDownColor = &HF0CDC0
   m_NormalButtonHoverColor = &HF09D90
   m_NormalButtonUpColor = &HF0CDC0
   m_NormalHeaderBackColor = &HFFFFFF
   m_NormalHeaderForeColor = &HC65D21
   m_NormalHeaderHoverColor = &HFF8E42
   m_NormalItemBackColor = &HF7DFD6
   m_NormalItemBorderColor = &HFFFFFF
   m_NormalItemForeColor = &HC65D21
   m_NormalItemHoverColor = &HFF8E42
   m_SpecialArrowDownColor = &HFF8E42
   m_SpecialArrowHoverColor = &HFFFFFF
   m_SpecialArrowUpColor = &HFFFFFF
   m_SpecialButtonBackColor = &HBC5215
   m_SpecialButtonDownColor = &HB24801
   m_SpecialButtonHoverColor = &HC6AD71
   m_SpecialButtonUpColor = &HC67D41
   m_SpecialHeaderBackColor = &HBC5215
   m_SpecialHeaderForeColor = &HFFFFFF
   m_SpecialHeaderHoverColor = &HFF8E42
   m_SpecialItemBackColor = &HF7DFD6
   m_SpecialItemBorderColor = &HFFFFFF
   m_SpecialItemForeColor = &HC65D21
   m_SpecialItemHoverColor = &HFF8E42
   m_UseAlphaBlend = True
   m_UseTheme = Windows
   picGroupMasker.Top = 0
   picGroupMasker.Left = 0
   PressedGroup = -1
   PressedItem = -1
   
   Call GetTextColors

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

   If (KeyCode <> vbKeyUp) And (KeyCode <> vbKeyDown) Or (KeyCode = vbKeySpace) Or (KeyCode = vbKeyReturn) Or tmrAnimation.Enabled Then Exit Sub
   
   PressedItem = -1
   PressedGroup = -1
   
   If FocussedItem > -1 Then
      If KeyCode = vbKeyUp Then
         FocussedItem = FocussedItem - 1
         
      ElseIf KeyCode = vbKeyDown Then
         FocussedItem = FocussedItem + 1
      End If
      
      If FocussedItem < 0 Then
         FocussedItem = -1
         
         If Groups(FocussedGroup).TypeGroup <> GROUP_DETAILS Then DrawItems FocussedGroup
         
         GoTo ExitSub
         
      ElseIf FocussedItem >= Groups(FocussedGroup).ItemsCount Then
         FocussedItem = -3
         
         If Groups(FocussedGroup).TypeGroup <> GROUP_DETAILS Then DrawItems FocussedGroup
         
         FocussedGroup = FocussedGroup + 1
         
         If FocussedGroup >= GroupsCount Then
            FocussedItem = -2
            FocussedGroup = GroupsCount - 1
         End If
         
      Else
         GoTo ExitSub
      End If
   End If
   
   If FocussedGroup > -1 Then
      If FocussedItem < -1 Then
         With Groups(FocussedGroup)
            If (.WindowState = Expanded) And (.TypeGroup <> GROUP_DETAILS) And Not .ContainerSet Then
               If (FocussedItem = -2) And .ItemsCount Then
                  FocussedItem = .ItemsCount - 1
                  
               Else
                  FocussedItem = -1
               End If
               
            Else
               FocussedItem = -1
            End If
         End With
         
      ElseIf KeyCode = vbKeyUp Then
         If FocussedItem = -1 Then
            DrawGroup FocussedGroup
            FocussedItem = -1
            FocussedGroup = FocussedGroup - 1
            
            If FocussedGroup < 0 Then
               FocussedGroup = 0
               
            Else
               With Groups(FocussedGroup)
                  If (.WindowState = Expanded) And (.TypeGroup <> GROUP_DETAILS) And Not .ContainerSet Then FocussedItem = .ItemsCount - 1
               End With
            End If
            
         Else
            FocussedGroup = 0
            FocussedItem = -1
         End If
         
      ElseIf KeyCode = vbKeyDown Then
         Call ResetFocussedGroup
         
         FocussedItem = -1
         
         With Groups(FocussedGroup)
            If (.WindowState = Expanded) And (.TypeGroup <> GROUP_DETAILS) And Not .ContainerSet Then
               If .ItemsCount Then FocussedItem = 0
               
            Else
               FocussedGroup = FocussedGroup + (1 And (FocussedGroup < GroupsCount - 1))
            End If
         End With
      End If
   End If
   
ExitSub:
   If FocussedGroup > -1 Then
      PressedGroup = FocussedGroup
      PressedItem = FocussedItem
      DrawGroup PressedGroup
      
      Call MakeGroupVisible(PressedGroup)
   End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

   If (KeyAscii <> vbKeySpace) And (KeyAscii <> vbKeyReturn) Or tmrAnimation.Enabled Then Exit Sub
   
   If FocussedItem > -1 Then
      KeyPressed = True
      PressedItem = FocussedItem
      
      Call picItems_MouseUp(FocussedGroup, vbLeftButton, 0, ITEM_SPACE * 3, (ITEM_SPACE + TextHeight("X") / 2) * (FocussedItem + 1))
      Call picItems_Click(FocussedGroup)
      
   ElseIf FocussedGroup > -1 Then
      KeyPressed = True
      PressedGroup = FocussedGroup
      
      Call picGroup_MouseUp(FocussedGroup, vbLeftButton, 0, ITEM_SPACE * 3, GROUP_SPACE / 2)
   End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Call ResetGroup

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Const SPI_GETWHEELSCROLLLINES As Long = 104

Dim blnTrackMouse             As Boolean

   With PropBag
      m_Animation = .ReadProperty("Animation", Medium)
      m_BackColor = .ReadProperty("BackColor", &HE6AA8C)
      bdrBorder.BorderColor = .ReadProperty("BorderColor", &HFFFFFF)
      m_DetailsForeColor = .ReadProperty("DetailsForeColor", &HC65D21)
      m_DetailGroupButton = .ReadProperty("DetailGroupButton", False)
      UserControl.Font = .ReadProperty("Font", UserControl.Font)
      UserControl.FontBold = False
      m_GradientBackColor = .ReadProperty("GradientBackColor", &HCC6633)
      m_GradientNormalHeaderBackColor = .ReadProperty("GradientNormalHeaderBackColor", &HF0D2C5)
      m_GradientNormalItemBackColor = .ReadProperty("GradientNormalItemBackColor", &HF7DFD6)
      m_GradientSpecialHeaderBackColor = .ReadProperty("GradientSpecialHeaderBackColor", &HB24801)
      m_GradientSpecialItemBackColor = .ReadProperty("GradientSpecialItemBackColor", &HF7DFD6)
      m_GradientStyle = .ReadProperty("GradientStyle", LeftRight)
      m_HeaderHeight = .ReadProperty("HeaderHeight", Low)
      m_Locked = .ReadProperty("Locked", False)
      m_NormalArrowDownColor = .ReadProperty("NormalArrowDownColor", &HC65D21)
      m_NormalArrowHoverColor = .ReadProperty("NormalArrowHoverColor", &HC65D21)
      m_NormalArrowUpColor = .ReadProperty("NormalArrowUpColor", &HFF8E42)
      m_NormalButtonBackColor = .ReadProperty("NormalButtonBackColor", &HFFFFFF)
      m_NormalButtonDownColor = .ReadProperty("NormalButtonDownColor", &HF0CDC0)
      m_NormalButtonHoverColor = .ReadProperty("NormalButtonHoverColor", &HF09D90)
      Set m_ButtonPicture(2) = .ReadProperty("NormalButtonPictureDown", Nothing)
      Set m_ButtonPicture(1) = .ReadProperty("NormalButtonPictureHover", Nothing)
      Set m_ButtonPicture(0) = .ReadProperty("NormalButtonPictureUp", Nothing)
      m_NormalButtonUpColor = .ReadProperty("NormalButtonUpColor", &HF0CDC0)
      m_NormalHeaderBackColor = .ReadProperty("NormalHeaderBackColor", &HFFFFFF)
      m_NormalHeaderForeColor = .ReadProperty("NormalHeaderForeColor", &HC65D21)
      m_NormalHeaderHoverColor = .ReadProperty("NormalHeaderHoverColor ", &HFF8E42)
      m_NormalItemBackColor = .ReadProperty("NormalItemBackColor", &HF7DFD6)
      m_NormalItemBorderColor = .ReadProperty("NormalItemBorderColor", &HFFFFFF)
      m_NormalItemForeColor = .ReadProperty("NormalItemForeColor", &HC65D21)
      m_NormalItemHoverColor = .ReadProperty("NormalItemHoverColor", &HFF8E42)
      m_OpenOneGroupOnly = .ReadProperty("OpenOneGroupOnly", False)
      bdrBorder.Visible = .ReadProperty("ShowBorder", False)
      m_SoundGroupClicked = .ReadProperty("SoundGroupClicked", False)
      m_SoundItemClicked = .ReadProperty("SoundItemClicked", False)
      m_SpecialArrowDownColor = .ReadProperty("SpecialArrowDownColor", &HFF8E42)
      m_SpecialArrowHoverColor = .ReadProperty("SpecialArrowHoverColor", &HFFFFFF)
      m_SpecialArrowUpColor = .ReadProperty("SpecialArrowUpColor", &HFFFFFF)
      m_SpecialButtonBackColor = .ReadProperty("SpecialButtonBackColor", &HBC5215)
      m_SpecialButtonDownColor = .ReadProperty("SpecialButtonDownColor", &HB24801)
      m_SpecialButtonHoverColor = .ReadProperty("SpecialButtonHoverColor", &HC6AD71)
      Set m_ButtonPicture(5) = .ReadProperty("SpecialButtonPictureDown", Nothing)
      Set m_ButtonPicture(4) = .ReadProperty("SpecialButtonPictureHover", Nothing)
      Set m_ButtonPicture(3) = .ReadProperty("SpecialButtonPictureUp", Nothing)
      m_SpecialButtonUpColor = .ReadProperty("SpecialButtonUpColor", &HC67D41)
      m_SpecialHeaderBackColor = .ReadProperty("SpecialHeaderBackColor", &HBC5215)
      m_SpecialHeaderForeColor = .ReadProperty("SpecialHeaderForeColor", &HFFFFFF)
      m_SpecialHeaderHoverColor = .ReadProperty("SpecialHeaderHoverColor", &HFF8E42)
      m_SpecialItemBackColor = .ReadProperty("SpecialItemBackColor", &HF7DFD6)
      m_SpecialItemBorderColor = .ReadProperty("SpecialItemBorderColor", &HFFFFFF)
      m_SpecialItemForeColor = .ReadProperty("SpecialItemForeColor", &HC65D21)
      m_SpecialItemHoverColor = .ReadProperty("SpecialItemHoverColor", &HFF8E42)
      m_UseAlphaBlend = .ReadProperty("UseAlphaBlend", True)
      m_UseTheme = .ReadProperty("UseTheme", Windows)
      m_UseUserForeColors = .ReadProperty("UseUserForeColors", False)
      UseButtonPictures = CountButtonPictures
      MoveLines = m_Animation * 3
   End With
   
   Call GetTextColors
   Call Refresh
   
   SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, ScrollLines, 0
   ScrollLines = ScrollLines + (1 And (ScrollLines = 0))
   
   If Ambient.UserMode Then
      Call SetAlignment
      
      TrackUser32 = IsFunctionSupported("TrackMouseEvent", "User32")
      
      If Not TrackUser32 Then blnTrackMouse = IsFunctionSupported("_TrackMouseEvent", "ComCtl32")
      
      With UserControl
         Subclass_Initialize .hWnd
         
         Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE)
         Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE)
         Call Subclass_AddMsg(.hWnd, WM_SYSCOLORCHANGE)
         Call Subclass_AddMsg(.hWnd, WM_THEMECHANGED)
      End With
      
      With tsbVertical
         Subclass_Initialize .hWnd
         
         Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE)
         Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE)
      End With
   End If

End Sub

Private Sub UserControl_Resize()

   If Width < 201 * Screen.TwipsPerPixelX Then Width = 201 * Screen.TwipsPerPixelX
   
   With tsbVertical
      .Left = ScaleWidth - (.Width And .Visible)
   End With
   
   With bdrBorder
      .Top = ScaleTop
      .Left = ScaleLeft
      .Width = ScaleWidth
      .Height = ScaleHeight
   End With
   
   If Not DoAlignment Then
      Call SetAlignment
      Call Refresh
      
   ElseIf DoAlignment And (Width > Extender.Parent.ScaleWidth) Then
      Width = Extender.Parent.ScaleWidth
      Extender.Align = vbAlignLeft
      
   Else
      Call Refresh
   End If

End Sub

Private Sub UserControl_Show()

   Call Refresh

End Sub

Private Sub UserControl_Terminate()

Dim intCount As Integer

   On Local Error GoTo ExitSub
   
   For intCount = 0 To GroupsCount - 1
      If Groups(intCount).ContainerSet Then
         SetParent Groups(intCount).Container.hWnd, Groups(intCount).Container.Parent.hWnd
         Set Groups(intCount).Container = Nothing
         Set picItems(intCount) = Nothing
         Exit For
      End If
   Next 'intCount
   
   For intCount = picGroup.Count - 1 To 1 Step -1
      Unload picGroup(intCount)
      Unload picItems(intCount)
   Next 'intCount
   
   For intCount = 0 To 5
      Set m_ButtonPicture(intCount) = Nothing
   Next 'intCount
   
   ReDim Groups(0) As GroupType
   
   Call Subclass_Terminate
   
ExitSub:
   On Local Error GoTo 0
   Erase SubclassData

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      UserControl.FontBold = False
      .WriteProperty "Animation", m_Animation, Medium
      .WriteProperty "BackColor", m_BackColor, &HE6AA8C
      .WriteProperty "BorderColor", bdrBorder.BorderColor, &HFFFFFF
      .WriteProperty "DetailsForeColor", m_DetailsForeColor, &HC65D21
      .WriteProperty "DetailGroupButton", m_DetailGroupButton, False
      .WriteProperty "Font", UserControl.Font
      .WriteProperty "GradientBackColor", m_GradientBackColor, &HCC6633
      .WriteProperty "GradientNormalHeaderBackColor", m_GradientNormalHeaderBackColor, &HF0D2C5
      .WriteProperty "GradientNormalItemBackColor", m_GradientNormalItemBackColor, &HF7DFD6
      .WriteProperty "GradientSpecialHeaderBackColor", m_GradientSpecialHeaderBackColor, &HB24801
      .WriteProperty "GradientSpecialItemBackColor", m_GradientSpecialItemBackColor, &HF7DFD6
      .WriteProperty "GradientStyle", m_GradientStyle, LeftRight
      .WriteProperty "HeaderHeight", m_HeaderHeight, Low
      .WriteProperty "Locked", m_Locked, False
      .WriteProperty "NormalArrowDownColor", m_NormalArrowDownColor, &HC65D21
      .WriteProperty "NormalArrowHoverColor", m_NormalArrowHoverColor, &HC65D21
      .WriteProperty "NormalArrowUpColor", m_NormalArrowUpColor, &HFF8E42
      .WriteProperty "NormalButtonBackColor", m_NormalButtonBackColor, &HFFFFFF
      .WriteProperty "NormalButtonDownColor", m_NormalButtonDownColor, &HF0CDC0
      .WriteProperty "NormalButtonHoverColor", m_NormalButtonHoverColor, &HF09D90
      .WriteProperty "NormalButtonPictureDown", m_ButtonPicture(2), Nothing
      .WriteProperty "NormalButtonPictureHover", m_ButtonPicture(1), Nothing
      .WriteProperty "NormalButtonPictureUp", m_ButtonPicture(0), Nothing
      .WriteProperty "NormalButtonUpColor", m_NormalButtonUpColor, &HF0CDC0
      .WriteProperty "NormalHeaderBackColor", m_NormalHeaderBackColor, &HFFFFFF
      .WriteProperty "NormalHeaderForeColor", m_NormalHeaderForeColor, &HC65D21
      .WriteProperty "NormalHeaderHoverColor", m_NormalHeaderHoverColor, &HFF8E42
      .WriteProperty "NormalItemBackColor", m_NormalItemBackColor, &HF7DFD6
      .WriteProperty "NormalItemBorderColor", m_NormalItemBorderColor, &HFFFFFF
      .WriteProperty "NormalItemForeColor", m_NormalItemForeColor, &HC65D21
      .WriteProperty "NormalItemHoverColor", m_NormalItemHoverColor, &HFF8E42
      .WriteProperty "OpenOneGroupOnly", m_OpenOneGroupOnly, False
      .WriteProperty "ShowBorder", bdrBorder.Visible, False
      .WriteProperty "SoundGroupClicked", m_SoundGroupClicked, False
      .WriteProperty "SoundItemClicked", m_SoundItemClicked, False
      .WriteProperty "SpecialArrowDownColor", m_SpecialArrowDownColor, &HFF8E42
      .WriteProperty "SpecialArrowHoverColor", m_SpecialArrowHoverColor, &HFFFFFF
      .WriteProperty "SpecialArrowUpColor", m_SpecialArrowUpColor, &HFFFFFF
      .WriteProperty "SpecialButtonBackColor", m_SpecialButtonBackColor, &HBC5215
      .WriteProperty "SpecialButtonDownColor", m_SpecialButtonDownColor, &HB24801
      .WriteProperty "SpecialButtonHoverColor", m_SpecialButtonHoverColor, &HC6AD71
      .WriteProperty "SpecialButtonPictureDown", m_ButtonPicture(5), Nothing
      .WriteProperty "SpecialButtonPictureHover", m_ButtonPicture(4), Nothing
      .WriteProperty "SpecialButtonPictureUp", m_ButtonPicture(3), Nothing
      .WriteProperty "SpecialButtonUpColor", m_SpecialButtonUpColor, &HC67D41
      .WriteProperty "SpecialHeaderBackColor", m_SpecialHeaderBackColor, &HBC5215
      .WriteProperty "SpecialHeaderForeColor", m_SpecialHeaderForeColor, &HFFFFFF
      .WriteProperty "SpecialHeaderHoverColor", m_SpecialHeaderHoverColor, &HFF8E42
      .WriteProperty "SpecialItemBackColor", m_SpecialItemBackColor, &HF7DFD6
      .WriteProperty "SpecialItemBorderColor", m_SpecialItemBorderColor, &HFFFFFF
      .WriteProperty "SpecialItemForeColor", m_SpecialItemForeColor, &HC65D21
      .WriteProperty "SpecialItemHoverColor", m_SpecialItemHoverColor, &HFF8E42
      .WriteProperty "UseAlphaBlend", m_UseAlphaBlend, True
      .WriteProperty "UseTheme", m_UseTheme, Windows
      .WriteProperty "UseUserForeColors", m_UseUserForeColors, False
   End With

End Sub

