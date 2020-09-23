VERSION 5.00
Begin VB.UserControl drawfield 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
End
Attribute VB_Name = "drawfield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim cb As Boolean

Dim i  As Integer
Dim plus1 As Boolean
Dim plus2 As Boolean
Dim plus3 As Boolean

Dim m_beginColor                As OLE_COLOR
Dim m_endColor                  As OLE_COLOR
Dim m_Value                     As Byte
Dim m_boxCount                  As Byte
Dim m_boxSpace                  As Byte

Const m_def_Value = 0
Const m_def_beginColor = &HFF
Const m_def_endColor = &HFF00
Const m_def_boxCount = 30
Const m_def_boxSpace = 2

Public Property Get boxCount() As Byte
  boxCount = m_boxCount
End Property

Public Property Let boxCount(ByVal New_boxCount As Byte)
  m_boxCount = New_boxCount
  If New_boxCount < 3 Then MsgBox "3-50": m_boxCount = 3
  If New_boxCount > 50 Then MsgBox "3-50": m_boxCount = 50
  PropertyChanged "boxCount"
End Property

Public Property Get boxSpace() As Byte
  boxSpace = m_boxSpace
End Property

Public Property Let boxSpace(ByVal New_boxSpace As Byte)
  m_boxSpace = New_boxSpace
  If New_boxSpace > 5 Then MsgBox "1-5": m_boxSpace = 5
  PropertyChanged "boxSpace"
End Property

Public Property Get Value() As Byte
  Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Byte)
  m_Value = New_Value
  If New_Value > 100 Then MsgBox "1-100": m_Value = 100
  PropertyChanged "Value"
  
  ncolor1 = Right$("000000" & Hex$(m_beginColor), 6)
  ncolor2 = Right$("000000" & Hex$(m_endColor), 6)
  ncolor1 = Mid$(ncolor1, 5, 2) + Mid$(ncolor1, 3, 2) + Mid$(ncolor1, 1, 2)
  ncolor2 = Mid$(ncolor2, 5, 2) + Mid$(ncolor2, 3, 2) + Mid$(ncolor2, 1, 2)
 
  Call draw(ncolor1, ncolor2, m_boxCount, m_boxSpace)
End Property


Public Property Get beginColor() As OLE_COLOR
  beginColor = m_beginColor
End Property

Public Property Let beginColor(ByVal New_beginColor As OLE_COLOR)
  m_beginColor = New_beginColor
  PropertyChanged "beginColor"
End Property

Public Property Get endColor() As OLE_COLOR
  endColor = m_endColor
End Property

Public Property Let endColor(ByVal New_endColor As OLE_COLOR)
  m_endColor = New_endColor
  PropertyChanged "endColor"
End Property

Private Sub UserControl_InitProperties()

i = 0: i2 = 0
m_beginColor = m_def_beginColor
m_endColor = m_def_endColor
m_Value = m_def_Value
m_boxCount = m_def_boxCount
m_boxSpace = m_def_boxSpace

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

m_beginColor = PropBag.ReadProperty("beginColor", m_def_beginColor)
m_endColor = PropBag.ReadProperty("endColor", m_def_endColor)
m_Value = PropBag.ReadProperty("Value", m_def_Value)
m_boxCount = PropBag.ReadProperty("boxCount", m_def_boxCount)
m_boxSpace = PropBag.ReadProperty("boxSpace", m_def_boxSpace)

End Sub

Private Sub UserControl_Resize()

If Width < 1000 Then Width = 1000
If Height < 200 Then Height = 200

Static IsR As Boolean
If IsR Then Exit Sub
IsR = True

If (Not m_boxCount = 0 And Not m_boxSpace = 0) Then
    dw = ScaleWidth
    
    Dim aw As Byte
    cnt = m_boxCount
    spa = m_boxSpace
    aw = ((dw - spa) / cnt)
    
    nw = (aw * cnt + 5)
    Width = nw * Screen.TwipsPerPixelX
End If
IsR = False

End Sub

Public Sub draw(cl1, cl2, cnt, space)

Dim color1 As String
Dim color2 As String
color1 = CStr(cl1)
color2 = CStr(cl2)
Dim aw As Byte
    
dw = ScaleWidth: aw = ((dw - space) / cnt)
s = (dw / 100 * Value) / aw: i = s: j = i
If i > cnt - 1 Then Exit Sub

'For j = 0 To i

X1 = space + j * aw
X2 = X1 + (aw - space)
Y1 = space - 1
Y2 = (ScaleHeight - space)

c11 = Val("&h" + Mid$(color1, 1, 2))
c12 = Val("&h" + Mid$(color1, 3, 2))
c13 = Val("&h" + Mid$(color1, 5, 2))
c21 = Val("&h" + Mid$(color2, 1, 2))
c22 = Val("&h" + Mid$(color2, 3, 2))
c23 = Val("&h" + Mid$(color2, 5, 2))

absC11C21_peraw = Int(Abs(c11 - c21) / cnt)
absC12C22_peraw = Int(Abs(c12 - c22) / cnt)
absC13C23_peraw = Int(Abs(c13 - c23) / cnt)

If c11 > c21 Then plus1 = True Else plus1 = False
If c12 > c22 Then plus2 = True Else plus2 = False
If c13 > c23 Then plus3 = True Else plus3 = False

If plus1 Then c31 = c11 - i * absC11C21_peraw
If plus2 Then c32 = c12 - i * absC12C22_peraw
If plus3 Then c33 = c13 - i * absC13C23_peraw
If Not plus1 Then c31 = c11 + i * absC11C21_peraw
If Not plus2 Then c32 = c12 + i * absC12C22_peraw
If Not plus3 Then c33 = c13 + i * absC13C23_peraw

If c31 <= 0 Then c31 = 0
If c32 <= 0 Then c32 = 0
If c33 <= 0 Then c33 = 0
If c31 >= 255 Then c31 = 255
If c32 >= 255 Then c32 = 255
If c33 >= 255 Then c33 = 255

Line (X1, Y1)-(X2, Y2), RGB(c31, c32, c33), BF
'Next j

For j = i + 1 To cnt
    X1 = space + j * aw
    X2 = X1 + (aw - space)
    Y1 = space - 1
    Y2 = (ScaleHeight - space)
    Line (X1, Y1)-(X2, Y2), RGB(255, 255, 255), BF
Next j
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Call PropBag.WriteProperty("beginColor", m_beginColor, m_def_beginColor)
Call PropBag.WriteProperty("endColor", m_endColor, m_def_endColor)
Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
Call PropBag.WriteProperty("boxCount", m_boxCount, m_def_boxCount)
Call PropBag.WriteProperty("boxSpace", m_boxSpace, m_def_boxSpace)

End Sub
