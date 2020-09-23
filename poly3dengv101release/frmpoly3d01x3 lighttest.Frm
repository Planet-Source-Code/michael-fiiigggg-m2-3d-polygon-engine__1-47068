VERSION 4.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   10590
   ClientLeft      =   3330
   ClientTop       =   2355
   ClientWidth     =   14880
   Height          =   10995
   Left            =   3270
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmpoly3d01x3 lighttest.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   706
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   992
   Top             =   2010
   Width           =   15000
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   10500
      Left            =   0
      ScaleHeight     =   696
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   896
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   13500
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public mousex As Integer, mousey As Integer
Public crss As Boolean
Public crrsi As Boolean
Public h6 As Variant
Public which2 As Variant
Public outlndd As Boolean
Public lightede As Boolean
Public maxdist As Integer
Public cfrrrammax As Integer
Public cleanmaxz As Integer

Private Sub Form_Activate()

'ok, messy, but more of "base - addons - addons - addons", so feh

'center form to screen at 70% screenheight/width
Form1.Top = (Screen.Height) * 0.15
Form1.Height = (Screen.Height * Screen.TwipsPerPixelY) ' * 0.5
Form1.Height = Form1.Height * 0.7
Form1.Left = (Screen.Width) * 0.15
Form1.Width = (Screen.Width * Screen.TwipsPerPixelX) '* 0.7
Form1.Width = Form1.Width * 0.7
'MsgBox Form1.Top / Screen.TwipsPerPixelY
Picture1.Left = 0
Picture1.Top = 0
Picture1.Width = Form1.Width / Screen.TwipsPerPixelX
Picture1.Height = Form1.Height / Screen.TwipsPerPixelY

'default for how many frames to clear
cleanmaxz = 50
'default for how many frames to advance last polygon
cfrrrammax = 20

crss = 0
DoEvents
outlndd = 1
DoEvents
OutlinedColor.R = 0
OutlinedColor.G = 0
OutlinedColor.b = 0
'Dim maxdist As Integer
lightede = True
maxdist = 40

ReDim Poly(0)     '!!!!!! need this for engine startup
Dim cd(9) As Double
'Dim h6 As Variant
Dim w3 As Integer
'Dim which2 As Variant

CameraPoint.X = 0
CameraPoint.Y = -60      '!!!!!! need this for engine startup
CameraPoint.z = 0
CameraAngle.X = 0
CameraAngle.Y = 0
CameraAngle.z = 0
CameraWindow.X = pi / 2      '!!!!!! need this for engine startup
CameraWindow.Y = pi / 3      '!!!!!! need this for engine startup
simpleYsort = True

h6 = InputBox("enter new number of polys")
'which2 = InputBox("enter new number of polys per object")
If h6 = "" Then
h6 = 1
End If
'If which2 = "" Then which2 = 2
'h6 = h6 * which2






'Dim Poly2 As T_3DPoly

ReDim Light(0)     '!!!!!! need this for engine startup
'Call AddLight(-3, -3, 1, RGB(200, 100, 200), 1, 10)
'Call AddLight(0, -60, 0, RGB(200, 100, 200), 1, 10)
'Call AddLight(0, -70, 0, RGB(0, 0, 255), 1, 1000, 150)
Call AddLight(0, -130, 0, RGB(255, 255, 255), 2, 1000, 200)
'Call AddLight(0, -30, 60, RGB(255, 255, 255), 1, 1000, 90)
'Call AddLight(0, -20, -100, RGB(255, 255, 255), 5, 1000, 90)

Dim fc, tmst

'Picture1.Cls
'Render3D Picture1
'BitBlt Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Picture1.hdc, 0, 0, SRCCOPY
'Form1.Refresh
DoEvents
fc = 0
tmst = Timer
Dim curpoly As Integer
curpoly = 0
'create 1st poly
Call Create3DPoly((RndNum(2) - 1) * 10, (RndNum(2) - 1) * 10, (RndNum(2) - 1) * 10 _
, (RndNum(2) - 1) * 10, (RndNum(2) - 1) * 10, (RndNum(2) - 1) * 10 _
, (RndNum(2) - 1) * 10, (RndNum(2) - 1) * 10, (RndNum(2) - 1) * 10 _
, RGB(20, 130, 100), lightede, RGB(0, 0, 0), RGB(255, 255, 255), outlndd)
Dim cfrrram As Integer
cfrrram = 1
Dim cleanct As Integer
cleanct = 1
Do
'the main loop for the 3d animation

'MsgBox Poly(1).p(2).x
cfrrram = cfrrram + 1
If cfrrram >= cfrrrammax Then   'if its time to advance polygon
'If RndNum(20) = 1 Then
cfrrram = 1
If UBound(Poly) < h6 Then     'if still creating polygons
'Form1.Caption = Poly(curpoly).p(2).x
'DoEvents


Call Create3DPoly(dbl(Poly(curpoly).p(2).X), dbl(Poly(curpoly).p(2).Y), dbl(Poly(curpoly).p(2).z), _
dbl(Poly(curpoly).p(3).X), dbl(Poly(curpoly).p(3).Y), dbl(Poly(curpoly).p(3).z), _
dbl(Poly(curpoly).p(3).X + ((RndNum(2) - 1) * 10)), dbl(Poly(curpoly).p(3).Y + ((RndNum(2) - 1) * 10)), dbl(Poly(curpoly).p(3).z + ((RndNum(2) - 1) * 10)), _
RGB(20, 130, 100), lightede, RGB(0, 0, 0), RGB(255, 255, 255), outlndd)
curpoly = curpoly + 1
If Abs(Poly(curpoly).p(1).X) > maxdist Then Poly(curpoly).p(1).X = (Poly(curpoly).p(1).X / Abs(Poly(curpoly).p(1).X)) * maxdist
If Abs(Poly(curpoly).p(1).Y) > maxdist Then Poly(curpoly).p(1).Y = (Poly(curpoly).p(1).Y / Abs(Poly(curpoly).p(1).Y)) * maxdist
If Abs(Poly(curpoly).p(1).z) > maxdist Then Poly(curpoly).p(1).z = (Poly(curpoly).p(1).z / Abs(Poly(curpoly).p(1).z)) * maxdist
If Abs(Poly(curpoly).p(2).X) > maxdist Then Poly(curpoly).p(2).X = (Poly(curpoly).p(2).X / Abs(Poly(curpoly).p(2).X)) * maxdist
If Abs(Poly(curpoly).p(2).Y) > maxdist Then Poly(curpoly).p(2).Y = (Poly(curpoly).p(2).Y / Abs(Poly(curpoly).p(2).Y)) * maxdist
If Abs(Poly(curpoly).p(2).z) > maxdist Then Poly(curpoly).p(2).z = (Poly(curpoly).p(2).z / Abs(Poly(curpoly).p(2).z)) * maxdist
If Abs(Poly(curpoly).p(3).X) > maxdist Then Poly(curpoly).p(3).X = (Poly(curpoly).p(3).X / Abs(Poly(curpoly).p(3).X)) * maxdist
If Abs(Poly(curpoly).p(3).Y) > maxdist Then Poly(curpoly).p(3).Y = (Poly(curpoly).p(3).Y / Abs(Poly(curpoly).p(3).Y)) * maxdist
If Abs(Poly(curpoly).p(3).z) > maxdist Then Poly(curpoly).p(3).z = (Poly(curpoly).p(3).z / Abs(Poly(curpoly).p(3).z)) * maxdist


Else  'if advancing already-made polygons
curpoly = curpoly + 1
If curpoly >= UBound(Poly) + 1 Then curpoly = 1
'Call Create3DPoly(Poly(curpoly).p(2).x, Poly(curpoly).p(2).y, Poly(curpoly).p(2).z, _
Poly(curpoly).p(3).x, Poly(curpoly).p(3).y, Poly(curpoly).p(3).z, _
Poly(curpoly).p(3).x + ((RndNum(2) - 1) * 10), Poly(curpoly).p(3).y + ((RndNum(2) - 1) * 10), Poly(curpoly).p(3).z + ((RndNum(2) - 1) * 10), _
RGB(20, 100, 100), True, RGB(5, 10, 10), RGB(255, 255, 255), True)
If curpoly = 1 Then

Poly(curpoly).p(1).X = Poly(UBound(Poly)).p(2).X
Poly(curpoly).p(1).Y = Poly(UBound(Poly)).p(2).Y
Poly(curpoly).p(1).z = Poly(UBound(Poly)).p(2).z
Poly(curpoly).p(2).X = Poly(UBound(Poly)).p(3).X
Poly(curpoly).p(2).Y = Poly(UBound(Poly)).p(3).Y
Poly(curpoly).p(2).z = Poly(UBound(Poly)).p(3).z
Poly(curpoly).p(3).X = Poly(UBound(Poly)).p(3).X + ((RndNum(4) - 2) * 5)
Poly(curpoly).p(3).Y = Poly(UBound(Poly)).p(3).Y + ((RndNum(4) - 2) * 5)
Poly(curpoly).p(3).z = Poly(UBound(Poly)).p(3).z + ((RndNum(4) - 2) * 5)
'something is messed up with the advancing i think, if you watch it, some of the points dont connect...
'but this program isnt super-important, and it works good enough, so i wont worry about it

Else
Poly(curpoly).p(1).X = Poly(curpoly - 1).p(2).X
Poly(curpoly).p(1).Y = Poly(curpoly - 1).p(2).Y
Poly(curpoly).p(1).z = Poly(curpoly - 1).p(2).z
Poly(curpoly).p(2).X = Poly(curpoly - 1).p(3).X
Poly(curpoly).p(2).Y = Poly(curpoly - 1).p(3).Y
Poly(curpoly).p(2).z = Poly(curpoly - 1).p(3).z
Poly(curpoly).p(3).X = Poly(curpoly - 1).p(3).X + ((RndNum(2) - 1) * 10)
Poly(curpoly).p(3).Y = Poly(curpoly - 1).p(3).Y + ((RndNum(2) - 1) * 10)
Poly(curpoly).p(3).z = Poly(curpoly - 1).p(3).z + ((RndNum(2) - 1) * 10)
End If

'keep all polys within maxdist [imaginary cube]
If Abs(Poly(curpoly).p(1).X) > maxdist Then Poly(curpoly).p(1).X = (Poly(curpoly).p(1).X / Abs(Poly(curpoly).p(1).X)) * maxdist
If Abs(Poly(curpoly).p(1).Y) > maxdist Then Poly(curpoly).p(1).Y = (Poly(curpoly).p(1).Y / Abs(Poly(curpoly).p(1).Y)) * maxdist
If Abs(Poly(curpoly).p(1).z) > maxdist Then Poly(curpoly).p(1).z = (Poly(curpoly).p(1).z / Abs(Poly(curpoly).p(1).z)) * maxdist
If Abs(Poly(curpoly).p(2).X) > maxdist Then Poly(curpoly).p(2).X = (Poly(curpoly).p(2).X / Abs(Poly(curpoly).p(2).X)) * maxdist
If Abs(Poly(curpoly).p(2).Y) > maxdist Then Poly(curpoly).p(2).Y = (Poly(curpoly).p(2).Y / Abs(Poly(curpoly).p(2).Y)) * maxdist
If Abs(Poly(curpoly).p(2).z) > maxdist Then Poly(curpoly).p(2).z = (Poly(curpoly).p(2).z / Abs(Poly(curpoly).p(2).z)) * maxdist
If Abs(Poly(curpoly).p(3).X) > maxdist Then Poly(curpoly).p(3).X = (Poly(curpoly).p(3).X / Abs(Poly(curpoly).p(3).X)) * maxdist
If Abs(Poly(curpoly).p(3).Y) > maxdist Then Poly(curpoly).p(3).Y = (Poly(curpoly).p(3).Y / Abs(Poly(curpoly).p(3).Y)) * maxdist
If Abs(Poly(curpoly).p(3).z) > maxdist Then Poly(curpoly).p(3).z = (Poly(curpoly).p(3).z / Abs(Poly(curpoly).p(3).z)) * maxdist

End If
End If
For s = 1 To UBound(Poly)
'rotate polys according to mouse position

'If Abs(mousex) > Abs(mousey) Then
Poly(s) = Rotate3DPoly(Poly(s), (pi / -5) * (mousey / Form1.ScaleHeight), 0, (pi / -5) * (mousex / Form1.ScaleWidth))
'Poly(s) = Rotate3DPoly(Poly(s), 0, 0, (pi / -60) * (mousex / Form1.ScaleWidth))
'Else
'Poly(s) = Rotate3DPoly(Poly(s), (pi / -60) * (mousey / Form1.ScaleHeight), 0, 0)
'End If
Next s
DoEvents
cleanct = cleanct + 1
If cleanct >= cleanmaxz Then cleanct = 1
If crss Or (crrsi And cleanct = 1) Then Picture1.Cls

', [pic1.cls,] render3d, bitblt, refresh are the needed parts for using render3d function
Render3D Picture1
BitBlt Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Picture1.hdc, 0, 0, SRCCOPY
'Form1.PaintPicture Picture1.Picture, 0, 0
Form1.Refresh
fc = fc + 1
DoEvents
If Timer - tmst >= 1 Then
Form1.Caption = fc & " fps   -functions: leftclick, rightclick, leftctrl, space, plus, minus, shift, up, down. cycle = " & cfrrrammax & ", clear = " & cleanmaxz
fc = 0
tmst = Timer
End If
'change light color, for world peace.
Light(1).colr.R = Light(1).colr.R + (RndNum(15) - RndNum(16))
Light(1).colr.G = Light(1).colr.G + (RndNum(15) - RndNum(16))
Light(1).colr.b = Light(1).colr.b + (RndNum(15) - RndNum(14))
If Light(1).colr.R < 0 Then Light(1).colr.R = 0
If Light(1).colr.R > 255 Then Light(1).colr.R = 255
If Light(1).colr.G < 0 Then Light(1).colr.G = 0
If Light(1).colr.G > 255 Then Light(1).colr.G = 255
If Light(1).colr.b < 0 Then Light(1).colr.b = 0
If Light(1).colr.b > 255 Then Light(1).colr.b = 255


Loop
'shizzle!

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'dis are de c00l shiz to do
If KeyCode = 17 Then
If outlndd = True Then
outlndd = False
Else
outlndd = True
End If
Dim p
For p = 1 To UBound(Poly)
Poly(p).outlined = outlndd
Next p
ElseIf KeyCode = 32 Then
If crrsi = True Then
crrsi = False
Else
crrsi = True
End If
ElseIf KeyCode = 16 Then
If lightede = True Then
lightede = False
Else
lightede = True
End If
For p = 1 To UBound(Poly)
Poly(p).lighted = lightede
Next p
ElseIf KeyCode = 40 Then
cfrrrammax = cfrrrammax + 1
ElseIf KeyCode = 38 Then
cfrrrammax = cfrrrammax - 1
If cfrrrammax <= 1 Then cfrrrammax = 1
ElseIf KeyCode = 189 Then
cleanmaxz = cleanmaxz - 1
If cleanmaxz <= 1 Then cleanmaxz = 1
ElseIf KeyCode = 187 Then
cleanmaxz = cleanmaxz + 1
End If
'MsgBox KeyCode
'189+ 187-
'38up40d

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'to reset num of polys, or toggle screen clearing
Form1.Caption = maxdist
If Button = 1 Then
ReDim Poly(0)
CameraPoint.Y = -80
Dim z As Integer
Dim cd(9) As Double
'Dim h6 As Variant
'Dim w3 As Integer
Dim which2 As Variant
h6 = InputBox("enter new number of polys")
'which2 = InputBox("enter new number of polys per object")
If h6 = "" Then
h6 = 3
End If
'If which2 = "" Then which2 = 5
'h6 = h6 * which2
Dim which As Integer
'''''''''''
'End If
For p = 2 To h6 + 1
For z = 1 To 9
DoEvents
cd(z) = (RndNum(maxdist / 1) - (maxdist / 2)) ' + (RndNum(maxdist) / 4 - maxdist / 2)
'MsgBox cd(z)
Next z
Call Create3DPoly(cd(1), cd(2), cd(3), cd(4), cd(5), cd(6), cd(7), cd(8), cd(9), RGB(20, 130, 100), lightede, RGB(5, 10, 10), RGB(255, 255, 255), outlndd)

Next p

ElseIf Button = 2 Then
If crss = True Then
crss = False
Else
crss = True
End If
End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousex = X - (Form1.ScaleWidth / 2)
mousey = Y - (Form1.ScaleHeight / 2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'since main program is in un-ending loop of form_activate, program would continue without end
End
End Sub


Private Sub Form_Resize()
'so the picturebox matches form


'Form1.Height = Screen.Height * Screen.TwipsPerPixelY
'Form1.Top = 0
'Form1.Left = 0
'Form1.Width = Screen.Width * Screen.TwipsPerPixelX
Picture1.Cls
Picture1.Left = 0
Picture1.Top = 0
Picture1.Width = Form1.Width / Screen.TwipsPerPixelX
Picture1.Height = Form1.Height / Screen.TwipsPerPixelY

End Sub


Private Sub Picture1_Click()

End Sub


