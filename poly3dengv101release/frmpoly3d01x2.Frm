VERSION 4.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   1245
   ClientTop       =   1485
   ClientWidth     =   9600
   Height          =   7545
   Left            =   1185
   LinkTopic       =   "Form1"
   ScaleHeight     =   476
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   Top             =   1140
   Width           =   9720
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7260
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public mousex As Integer, mousey As Integer

Private Sub Form_Activate()

DoEvents
'Form_CenterScreen Picture1
DoEvents


  'draw in picbox ==========================
'Picture1.Cls
'DrawTriangle Picture1, Coords(), Int(RGB(0, 255, 0))
  'bitblt to form ===========================
'BitBlt Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY
'Form1.Refresh
ReDim Poly(0)
'Call Create3DPoly(-10, -4, -6, -12, -9, -8, -15, -4, -2, RGB(0, 255, 255))
'error-Call Create3DPoly(3, 4, 2, 7, 3, 5, 7, 8, 4, RGB(0, 0, 255))
'Poly(2) = Rotate3DPoly(Poly(1), 0, pi / 4, 0)
CameraPoint.X = 0
CameraPoint.Y = -4
CameraPoint.z = 0
CameraAngle.X = 0
CameraAngle.Y = 0
CameraAngle.z = 0
CameraWindow.X = pi / 2
CameraWindow.Y = pi / 3









Dim Poly2 As T_3DPoly

ReDim Light(0)
'Call AddLight(-3, -3, 1, RGB(200, 100, 200), 1, 10)
'Call AddLight(0, -60, 0, RGB(200, 100, 200), 1, 10)
Call AddLight(-60, 0, 0, RGB(200, 100, 200), 1, 100, 10)
'Call Create3DPoly(-0.1, -2.9, -0.1, 0, -3, 0.1, 0.1, -3.1, -0.1, RGB(255, 255, 0))

Call Create3DPoly(-1, -1, 1, -1, -1, -1, 1, -1, 1, RGB(255, 255, 0), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
Call Create3DPoly(1, -1, -1, -1, -1, -1, 1, -1, 1, RGB(255, 255, 0), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
Call Create3DPoly(-1, 1, 1, -1, 1, -1, 1, 1, 1, RGB(0, 255, 0), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
Call Create3DPoly(1, 1, -1, -1, 1, -1, 1, 1, 1, RGB(0, 255, 0), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
Call Create3DPoly(-1, -1, 1, -1, -1, -1, -1, 1, 1, RGB(255, 0, 0), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
Call Create3DPoly(-1, -1, -1, -1, 1, -1, -1, 1, 1, RGB(255, 0, 0), True, RGB(100, 0, 0), RGB(255, 255, 255), False)

Call Create3DPoly(1, -1, 1, 1, -1, -1, 1, 1, 1, RGB(255, 255, 255), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
'Poly2 = Poly(1)

Call Create3DPoly(1, -1, -1, 1, 1, -1, 1, 1, 1, RGB(255, 255, 255), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
Call Create3DPoly(-1, -1, 1, 1, -1, 1, 1, 1, 1, RGB(255, 255, 255), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
Call Create3DPoly(-1, -1, 1, -1, 1, 1, 1, 1, 1, RGB(255, 0, 255), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
Call Create3DPoly(-1, -1, -1, 1, -1, -1, 1, 1, -1, RGB(0, 255, 255), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
Call Create3DPoly(-1, -1, -1, -1, 1, -1, 1, 1, -1, RGB(0, 255, 255), True, RGB(100, 0, 0), RGB(255, 255, 255), False)
Form_Click


Dim fc, tmst


Picture1.Cls
Render3D Picture1
BitBlt Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY
Form1.Refresh
DoEvents
fc = 0
tmst = Timer
Do
'Pause 0.1
For s = 1 To UBound(Poly)
'If Abs(mousex) > Abs(mousey) Then
Poly(s) = Rotate3DPoly(Poly(s), (pi / -5) * (mousey / Form1.ScaleHeight), 0, (pi / -5) * (mousex / Form1.ScaleWidth))
'Poly(s) = Rotate3DPoly(Poly(s), 0, 0, (pi / -60) * (mousex / Form1.ScaleWidth))
'Else
'Poly(s) = Rotate3DPoly(Poly(s), (pi / -60) * (mousey / Form1.ScaleHeight), 0, 0)
'End If
Next s
DoEvents
Picture1.Cls
Render3D Picture1
BitBlt Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY
Form1.Refresh
fc = fc + 1
DoEvents
If Timer - tmst >= 1 Then
Form1.Caption = fc & " fps"
fc = 0
tmst = Timer
End If

Loop



End Sub


Private Sub Form_Click()
ReDim Poly(0)
CameraPoint.Y = -100

Dim cd(9) As Double
Dim h6 As Variant
Dim w3 As Integer
Dim which2 As Variant
h6 = InputBox("enter new number of objects")
which2 = InputBox("enter new number of polys per object")
If h6 = "" Then
h6 = 3
End If
If which2 = "" Then which2 = 5
h6 = h6 * which2
Call Create3DPoly(-10, -10, 10, -10, -10, -10, 10, -10, 10, RGB(0, 0, 255), False, RGB(100, 0, 255), RGB(255, 255, 255), False)
Call Create3DPoly(10, -10, -10, -10, -10, -10, 10, -10, 10, RGB(0, 0, 255), False, RGB(100, 0, 255), RGB(255, 255, 255), False)

Call Create3DPoly(-10, 10, 10, -10, 10, -10, 10, 10, 10, RGB(0, 255, 255), False, RGB(100, 0, 255), RGB(255, 255, 255), False)
Call Create3DPoly(10, 10, -10, -10, 10, -10, 10, 10, 10, RGB(0, 0, 255), False, RGB(100, 0, 255), RGB(255, 255, 255), False)
Dim which As Integer

cd(1) = RndNum(60) - 30
cd(2) = RndNum(60) - 30
cd(3) = RndNum(60) - 30
cd(4) = RndNum(60) - 30
cd(5) = RndNum(60) - 30
cd(6) = RndNum(60) - 30
cd(7) = RndNum(60) - 30
cd(8) = RndNum(60) - 30
cd(9) = RndNum(60) - 30
which = 0
w3 = 0
For p = 2 To h6 + 1
which = which + 1

If which = 4 Then which = 1
Select Case which
Case 1
cd(1) = cd(4) + RndNum(40) - 20
cd(2) = cd(5) + RndNum(40) - 20
cd(3) = cd(6) + RndNum(40) - 20
Case 2
cd(4) = cd(7) + RndNum(40) - 20
cd(5) = cd(8) + RndNum(40) - 20
cd(6) = cd(9) + RndNum(40) - 20
Case 3
cd(7) = cd(1) + RndNum(40) - 20
cd(8) = cd(2) + RndNum(40) - 20
cd(9) = cd(3) + RndNum(40) - 20
End Select
w3 = w3 + 1
If w3 = which2 Then
w3 = 0
cd(1) = RndNum(60) - 30
cd(2) = RndNum(60) - 30
cd(3) = RndNum(60) - 30
cd(4) = RndNum(60) - 30
cd(5) = RndNum(60) - 30
cd(6) = RndNum(60) - 30
cd(7) = RndNum(60) - 30
cd(8) = RndNum(60) - 30
cd(9) = RndNum(60) - 30
End If
'AD cd(1) & " " & cd(2) & " " & cd(3) & ", " & cd(4) & " " & cd(5) & " " & cd(6) & ", " & cd(7) & " " & cd(8) & " " & cd(9)
'Call Create3DPoly(cd(1), cd(2), cd(3), cd(4), cd(5), cd(6), cd(7), cd(8), cd(9), RGB(((p * 30) Mod 255), ((p * 30) Mod 255), ((p * 30) Mod 255))) 'RGB(RndNum(255), RndNum(255), RndNum(255)))
Call Create3DPoly(cd(1), cd(2), cd(3), cd(4), cd(5), cd(6), cd(7), cd(8), cd(9), RGB(0, 0, 200), True, RGB(0, 0, 100), RGB(255, 255, 255), False) 'RGB(((255 - 30) * ((p - 1) / h6)) + 30, ((255 - 30) * ((p - 1) / h6)) + 30, ((255 - 30) * ((p - 1) / h6)) + 30))
''Picture1.Cls
'Render3D Picture1
'BitBlt Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY
'Form1.Refresh
'Pause 1.2
Next p
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousex = X - (Form1.ScaleWidth / 2)
mousey = Y - (Form1.ScaleHeight / 2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


