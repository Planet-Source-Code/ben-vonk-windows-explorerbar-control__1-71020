VERSION 5.00
Begin VB.Form frmExplorerBar 
   Caption         =   "ExplorerBar Sample - by Ben Vonk"
   ClientHeight    =   6252
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   10656
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExplorerBar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6252
   ScaleWidth      =   10656
   StartUpPosition =   2  'CenterScreen
   Begin prjExplorerBar.ExplorerBar exbSideBarUser 
      Align           =   4  'Align Right
      Height          =   6252
      Left            =   7644
      TabIndex        =   14
      Top             =   0
      Width           =   3012
      _ExtentX        =   5313
      _ExtentY        =   11028
      BorderColor     =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientStyle   =   0
      UseTheme        =   2
      UseUserForeColors=   -1  'True
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   2652
      Left            =   3120
      ScaleHeight     =   2652
      ScaleWidth      =   2412
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2412
      Begin VB.CommandButton Command3 
         Caption         =   "&Clear"
         Height          =   372
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   852
      End
      Begin VB.Frame Frame1 
         Caption         =   "Demo:"
         Height          =   972
         Left            =   1080
         TabIndex        =   10
         Top             =   1560
         Width           =   1212
         Begin VB.CheckBox Check1 
            Caption         =   "Demo"
            Height          =   252
            Left            =   120
            TabIndex        =   13
            Top             =   660
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Demo"
            Height          =   252
            Left            =   120
            TabIndex        =   12
            Top             =   420
            Width           =   972
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Demo"
            Height          =   252
            Left            =   120
            TabIndex        =   11
            Top             =   180
            Width           =   972
         End
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   1080
         TabIndex        =   5
         Top             =   120
         Width           =   1212
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   372
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   852
      End
      Begin VB.ListBox lstResult 
         Height          =   552
         ItemData        =   "frmExplorerBar.frx":08CA
         Left            =   120
         List            =   "frmExplorerBar.frx":08CC
         TabIndex        =   7
         Top             =   720
         Width           =   2160
      End
      Begin VB.Label lblInfo 
         Caption         =   "Search:"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   132
         Width           =   852
      End
      Begin VB.Label lblInfo 
         Caption         =   "Result:"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   852
      End
   End
   Begin VB.PictureBox picEvents 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   3120
      ScaleHeight     =   372
      ScaleWidth      =   372
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3360
      Width           =   372
   End
   Begin VB.ListBox lstEvents 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      ItemData        =   "frmExplorerBar.frx":08CE
      Left            =   3120
      List            =   "frmExplorerBar.frx":08D0
      TabIndex        =   2
      Top             =   3840
      Width           =   4452
   End
   Begin prjExplorerBar.ExplorerBar exbSideBarThemed 
      Align           =   3  'Align Left
      Height          =   6252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3012
      _ExtentX        =   5313
      _ExtentY        =   11028
      DetailGroupButton=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SoundGroupClicked=   -1  'True
      SoundItemClicked=   -1  'True
      UseTheme        =   1
   End
   Begin VB.Image imgShare 
      Height          =   192
      Left            =   6720
      Picture         =   "frmExplorerBar.frx":08D2
      Top             =   5640
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgUpload 
      Height          =   192
      Left            =   7440
      Picture         =   "frmExplorerBar.frx":0E5C
      Top             =   5640
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgNew 
      Height          =   192
      Left            =   6720
      Picture         =   "frmExplorerBar.frx":13E6
      Top             =   5400
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgBurn 
      Height          =   192
      Left            =   7080
      Picture         =   "frmExplorerBar.frx":1970
      Top             =   5640
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgSlide 
      Height          =   192
      Left            =   7800
      Picture         =   "frmExplorerBar.frx":1EFA
      Top             =   5400
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgOrder 
      Height          =   192
      Left            =   7080
      Picture         =   "frmExplorerBar.frx":2484
      Top             =   5400
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgImages 
      Height          =   384
      Left            =   6600
      Picture         =   "frmExplorerBar.frx":2A0E
      Top             =   4920
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgCan 
      Height          =   4800
      Left            =   5040
      Picture         =   "frmExplorerBar.frx":5ED8
      Top             =   120
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Image imgBack 
      Height          =   648
      Left            =   5760
      Picture         =   "frmExplorerBar.frx":A80D
      Top             =   4920
      Visible         =   0   'False
      Width           =   684
   End
   Begin VB.Image imgPrint 
      Height          =   192
      Left            =   7440
      Picture         =   "frmExplorerBar.frx":CC97
      Top             =   5400
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgClick 
      Height          =   384
      Left            =   7080
      Picture         =   "frmExplorerBar.frx":D221
      Top             =   4920
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmExplorerBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SelectedGroup As Integer

Private Function GetObject(ByVal Group As Integer, ByVal Item As Integer) As String

Dim strObject As String

   If Group = -1 Then
      strObject = "ExplorerBar Control"
      
   ElseIf Group = -2 Then
      strObject = "ScrollBar"
      
   Else
      strObject = "Group: " & Group
      
      If Item > -1 Then strObject = strObject & " - Item: " & Item
   End If
   
   GetObject = strObject

End Function

Private Sub SetEventInfo(ByVal EventInfo As String, Optional ByRef EventImage As StdPicture = Nothing)

   With lstEvents
      If .ListCount = 50 Then .RemoveItem 0
      
      .AddItem EventInfo
      .ListIndex = .ListCount - 1
   End With
   
   picEvents.Picture = EventImage

End Sub

Private Sub exbSideBarThemed_GroupClick(Group As Integer, WindowState As WindowStates)

Dim strWindowState As String

   Select Case WindowState
      Case Expanded
         strWindowState = "Expanded]"
         
      Case Collapsed
         strWindowState = "Collapsed]"
         
      Case Fixed
         strWindowState = "Fixed]"
   End Select
   
   Call SetEventInfo("GroupClick " & Group & " (" & exbSideBarThemed.GetGroupTitle(Group) & ") [State: " & strWindowState, exbSideBarThemed.GetGroupIcon(Group))
   
   SelectedGroup = Group

End Sub

Private Sub exbSideBarThemed_ItemClick(Group As Integer, Item As Integer)

   Call SetEventInfo("ItemClick: " & Item & " in Group: " & Group & " (" & exbSideBarThemed.GetItemCaption(Group, Item) & ")", exbSideBarThemed.GetItemIcon(Group, Item))

End Sub

Private Sub exbSideBarThemed_ItemOpenFile(Group As Integer, Item As Integer, File As String)

   Call SetEventInfo("OpenFile for Item: " & Item & " in Group: " & Group & " (" & File & ")", exbSideBarThemed.GetItemIcon(Group, Item))

End Sub

Private Sub exbSideBarThemed_MouseHover(Group As Integer, Item As Integer, FullTextShowed As Boolean)

   Call SetEventInfo("MouseHover: " & GetObject(Group, Item))

End Sub

Private Sub exbSideBarThemed_MouseOut(Group As Integer, Item As Integer)

   Call SetEventInfo("MouseOut: " & GetObject(Group, Item))

End Sub

Private Sub Form_Load()

Dim intGroup As Integer

   SelectedGroup = -1
   
   With exbSideBarThemed
      .Locked = True
      intGroup = .AddNormalGroup("Container", , Expanded, , , , , picContainer)
      intGroup = .AddNormalGroup("Foldertasks", , Expanded)
      .AddItem intGroup, "New folder", , imgNew.Picture
      .AddItem intGroup, "Upload folder", , imgUpload.Picture
      .AddItem intGroup, "Share folder", , imgShare.Picture
      intGroup = .AddSpecialGroup("Imagetasks", , , imgImages.Picture, imgBack.Picture)
      .AddItem intGroup, "Show only text", True, imgSlide.Picture, True
      .AddItem intGroup, "Paint", , imgOrder.Picture, , "MSPaint.exe"
      .AddItem intGroup, "Print", , imgPrint.Picture
      .AddItem intGroup, "Planet Source Code", , imgBurn.Picture, , "http://www.planet-source-code.com/vb/default.asp?lngWId=1"
      intGroup = .AddNormalGroup("Other foldertasks", True, Collapsed)
      .AddItem intGroup, "New folder", , imgNew.Picture
      .AddItem intGroup, "Upload folder", , imgUpload.Picture
      .AddItem intGroup, "Share folder", , imgShare.Picture
      intGroup = .AddDetailGroup("More details", , Fixed, , imgCan.Picture, "Bthundertaste.jpg", "JPG-image" & vbCrLf & "Dimensions: 400 x 400" & vbCrLf & "Size: 18,2 kB" & vbCrLf & "Last modified: jul-2-2005 11:13")
      intGroup = .AddNormalGroup("Group with columns", , Collapsed)
      .AddItem intGroup, "New" & vbTab & "folder", True
      .AddItem intGroup, "Upload" & vbTab & "folder"
      .AddItem intGroup, "Share" & vbTab & "folder", , imgShare.Picture
      .AddItem intGroup, "Download folder", , imgUpload.Picture
      .Locked = False
   End With
   
   With exbSideBarUser
      .Locked = True
      .GradientStyle = RightLeft
      intGroup = .AddSpecialGroup("Imagetasks", , , imgImages.Picture, imgBack.Picture)
      .AddItem intGroup, "Make slideshow", , imgSlide.Picture
      .AddItem intGroup, "Order print online", , imgOrder.Picture
      .AddItem intGroup, "Print", , imgPrint.Picture
      .AddItem intGroup, "Burn on cd", , imgBurn.Picture
      intGroup = .AddNormalGroup("Foldertasks", , Collapsed)
      .AddItem intGroup, "New folder", , imgNew.Picture
      .AddItem intGroup, "Upload folder", , imgUpload.Picture
      .AddItem intGroup, "Share folder", , imgShare.Picture
      intGroup = .AddNormalGroup("Other foldertasks", , Collapsed)
      .AddItem intGroup, "New folder", , imgNew.Picture
      .AddItem intGroup, "Upload folder", , imgUpload.Picture
      .AddItem intGroup, "Share folder", , imgShare.Picture
      intGroup = .AddNormalGroup("Group with columns", , Collapsed)
      .AddItem intGroup, "New" & vbTab & "folder", True
      .AddItem intGroup, "Upload" & vbTab & "folder"
      .AddItem intGroup, "Share" & vbTab & "folder", , imgShare.Picture
      .AddItem intGroup, "Download folder", , imgUpload.Picture
      intGroup = .AddDetailGroup("More details", , Fixed, , imgCan.Picture, "thundertaste.jpg", "JPG-image" & vbCrLf & "Dimensions: 400 x 400" & vbCrLf & "Size: 18,2 kB" & vbCrLf & "Last modified: jul-2-2005 11:13")
      .Locked = False
   End With

End Sub

