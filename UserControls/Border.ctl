VERSION 5.00
Begin VB.UserControl Border 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   468
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   816
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   ScaleHeight     =   39
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   68
   ToolboxBitmap   =   "Border.ctx":0000
End
Attribute VB_Name = "Border"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Border Control
'
'Author Ben Vonk
'05-03-2008 First version

Option Explicit

' Public Enumeration
Public Enum BorderWidths
   Thin
   Thick
End Enum

' Private Variables
Private m_RoundBorder As Boolean
Private m_BorderWidth As BorderWidths
Private m_BorderColor As OLE_COLOR

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."

   BorderColor = m_BorderColor

End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)

   m_BorderColor = NewBorderColor
   PropertyChanged "BorderColor"
   
   Call Refresh

End Property

Public Property Get BorderWidth() As BorderWidths
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."

   BorderWidth = m_BorderWidth

End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As BorderWidths)

   m_BorderWidth = NewBorderWidth
   PropertyChanged "BorderWidth"
   
   Call Refresh

End Property

Public Property Get RoundBorder() As Boolean
Attribute RoundBorder.VB_Description = "Returns/sets a value that determines whether the border will be rounded."

   RoundBorder = m_RoundBorder

End Property

Public Property Let RoundBorder(ByVal NewRoundBorder As Boolean)

   m_RoundBorder = NewRoundBorder
   PropertyChanged "RoundBorder"
   
   Call Refresh

End Property

Public Sub Refresh()

Dim blnAutoRedraw As Boolean
Dim lngLeft       As Long
Dim lngHeight     As Long
Dim lngTop        As Long
Dim lngWidth      As Long

   With Extender
      lngTop = .Top
      lngLeft = .Left
      lngWidth = .Width
      lngHeight = .Height
   End With
   
   Cls
   DrawWidth = m_BorderWidth + 1
   Line (m_BorderWidth, m_BorderWidth)-(ScaleWidth - 1, ScaleHeight - 1), m_BorderColor, B
   MaskPicture = Image

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

   Call Refresh

End Sub

Private Sub UserControl_Initialize()

   m_BorderColor = vbWindowText
   m_RoundBorder = True

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      m_BorderColor = .ReadProperty("BorderColor", vbWindowText)
      m_BorderWidth = .ReadProperty("BorderWidth", Thin)
      m_RoundBorder = .ReadProperty("RoundBorder", True)
   End With
   
   Call Refresh

End Sub

Private Sub UserControl_Resize()

   Call Refresh

End Sub

Private Sub UserControl_Show()

   Call Refresh

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "BorderColor", m_BorderColor, vbBlack
      .WriteProperty "BorderWidth", m_BorderWidth, Thin
      .WriteProperty "RoundBorder", m_RoundBorder, True
   End With

End Sub

