Attribute VB_Name = "Module1"
Type T_3DPoly
p(3) As T_3DPoint
colr As T_RGB
colrmin As T_RGB
colrmax As T_RGB
lighted As Boolean
outlined As Boolean
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type


Public Poly() As T_3DPoly 'storage for all polygons' coords and colors.
Public CameraPoint As T_3DPoint 'camera position
Public CameraAngle As T_3DPoint 'camera angle
Public CameraWindow As T_3DPoint '"lens size"?  view window.
Public OutlinedColor As T_RGB
Public simpleYsort As Boolean

Type T_Light

  p As T_3DPoint
  colr As T_RGB
  val As Double
  dist As Double
  halflife As Double
  
End Type

Public Light() As T_Light

Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long


Sub AddLight(X, Y, z, colr, val, dist, halflife)
'adds a light.  val is default1, dist is the max the light shines, halflife is how far light gets before it's half as bright.
ReDim Preserve Light(UBound(Light) + 1)
Light(UBound(Light)).p.X = X
Light(UBound(Light)).p.Y = Y
Light(UBound(Light)).p.z = z
Light(UBound(Light)).colr.R = GetRGB(colr, "R")
Light(UBound(Light)).colr.G = GetRGB(colr, "G")
Light(UBound(Light)).colr.b = GetRGB(colr, "B")
Light(UBound(Light)).val = val
Light(UBound(Light)).dist = dist
Light(UBound(Light)).halflife = halflife


End Sub


Sub Create3DPoly(x1 As Double, y1 As Double, z1 As Double, x2 As Double, _
y2 As Double, z2 As Double, x3 As Double, y3 As Double, z3 As Double, colr As Long, _
lighted As Boolean, colrmin As Long, colrmax As Long, outlined As Boolean)
'creates 3d polygon using coords and color as parameters

ReDim Preserve Poly(UBound(Poly) + 1)
Poly(UBound(Poly)).p(1).X = x1
Poly(UBound(Poly)).p(1).Y = y1
Poly(UBound(Poly)).p(1).z = z1
Poly(UBound(Poly)).p(2).X = x2
Poly(UBound(Poly)).p(2).Y = y2
Poly(UBound(Poly)).p(2).z = z2
Poly(UBound(Poly)).p(3).X = x3
Poly(UBound(Poly)).p(3).Y = y3
Poly(UBound(Poly)).p(3).z = z3

Poly(UBound(Poly)).colr.R = GetRGB(colr, "R")
Poly(UBound(Poly)).colr.G = GetRGB(colr, "G")
Poly(UBound(Poly)).colr.b = GetRGB(colr, "B")

Poly(UBound(Poly)).lighted = lighted
Poly(UBound(Poly)).colrmin.R = GetRGB(colrmin, "R")
Poly(UBound(Poly)).colrmin.G = GetRGB(colrmin, "G")
Poly(UBound(Poly)).colrmin.b = GetRGB(colrmin, "B")

Poly(UBound(Poly)).colrmax.R = GetRGB(colrmax, "R")
Poly(UBound(Poly)).colrmax.G = GetRGB(colrmax, "G")
Poly(UBound(Poly)).colrmax.b = GetRGB(colrmax, "B")


Poly(UBound(Poly)).outlined = outlined


End Sub


Function dbl(X)
'just for trying something that wouldn't work without this quickfix thing
Dim Y As Double
Y = X
dbl = Y
End Function

Sub drawtri3(where As Object, Coords() As T_Point, colr As Long, outl As Boolean)
'draws triangle on screen.
Dim X(3) As POINTAPI
X(0).X = Int(Coords(0).X)
X(0).Y = Int(Coords(0).Y)
X(1).X = Int(Coords(1).X)
X(1).Y = Int(Coords(1).Y)
X(2).X = Int(Coords(2).X)
X(2).Y = Int(Coords(2).Y)
If outl = True Then
where.ForeColor = RGB(OutlinedColor.R, OutlinedColor.G, OutlinedColor.b)
Else
where.ForeColor = colr
End If
where.FillColor = colr
Polygon where.hdc, X(0), 3
End Sub


Sub DrawTriangle2(where As Object, Coords() As T_Point, colr As Long)
'draws outline of triangle, but not needed anymore
where.Line (Coords(0).X, Coords(0).Y)-(Coords(1).X, Coords(1).Y), colr
where.Line (Coords(1).X, Coords(1).Y)-(Coords(2).X, Coords(2).Y), colr
where.Line (Coords(2).X, Coords(2).Y)-(Coords(0).X, Coords(0).Y), colr

'where.Line (Coords(0).X - 10, Coords(0).Y)-(Coords(0).X + 3, Coords(0).Y + 10), RGB(255, 255, 255)
'where.Line (Coords(0).X + 3, Coords(0).Y + 10)-(Coords(0).X + 3, Coords(0).Y - 10), RGB(255, 255, 255)
'where.Line (Coords(0).X + 3, Coords(0).Y - 10)-(Coords(0).X - 10, Coords(0).Y), RGB(255, 255, 255)

'where.Line (Coords(1).X - 10, Coords(1).Y)-(Coords(1).X + 3, Coords(1).Y + 10), RGB(160, 160, 160)
'where.Line (Coords(1).X + 3, Coords(1).Y + 10)-(Coords(1).X + 3, Coords(1).Y - 10), RGB(160, 160, 160)
'where.Line (Coords(1).X + 3, Coords(1).Y - 10)-(Coords(1).X - 10, Coords(1).Y), RGB(160, 160, 160)

'where.Line (Coords(2).X - 10, Coords(2).Y)-(Coords(2).X + 3, Coords(2).Y + 10), RGB(70, 70, 70)
'where.Line (Coords(2).X + 3, Coords(2).Y + 10)-(Coords(2).X + 3, Coords(2).Y - 10), RGB(70, 70, 70)
'where.Line (Coords(2).X + 3, Coords(2).Y - 10)-(Coords(2).X - 10, Coords(2).Y), RGB(70, 70, 70)

End Sub

Sub Render3D(where As Object)
'renders the scene using lights, camera, and polygons
'main function of 3d engine.  some lines of code commented, should probably not be de-commented.
'basically, to use engine, in a form, redim light(0), redim poly(0), set the camera position and viewing window, and call this function over and over =p
Dim cyn As Integer
Dim NewPos As T_3DPoly
Dim ScrPos(3) As T_Point  'use 0,1,2 for this  =[  change later?
Dim Slope As Double
Dim RotAngle As T_3DPoint

Dim polylightdist As T_3DPoint


Dim Ordrd() As T_Point

ReDim Ordrd(UBound(Poly))

'yeah, weird names used in this engine...  i was just in need of a name, so sue me.

'this part sorts the polys by distance from camera... its not the 100% correct way, but it does enough for my satisfaction.  surprisingly doesnt take much time.
Dim pie
For pie = 1 To UBound(Poly)
Ordrd(pie).X = pie
'Ordrd(pie).Y = (Poly(pie).p(1).Y + Poly(pie).p(2).Y + Poly(pie).p(3).Y) / 3
If simpleYsort = False Then
NewPos = Poly(pie)
NewPos.p(1).X = NewPos.p(1).X - CameraPoint.X
NewPos.p(1).Y = NewPos.p(1).Y - CameraPoint.Y
NewPos.p(1).z = NewPos.p(1).z - CameraPoint.z

NewPos.p(2).X = NewPos.p(2).X - CameraPoint.X
NewPos.p(2).Y = NewPos.p(2).Y - CameraPoint.Y
NewPos.p(2).z = NewPos.p(2).z - CameraPoint.z

NewPos.p(3).X = NewPos.p(3).X - CameraPoint.X
NewPos.p(3).Y = NewPos.p(3).Y - CameraPoint.Y
NewPos.p(3).z = NewPos.p(3).z - CameraPoint.z

NewPos = Rotate3DPoly(NewPos, CameraAngle.X, CameraAngle.Y, CameraAngle.z)

Ordrd(pie).Y = (NewPos.p(1).Y + NewPos.p(2).Y + NewPos.p(3).Y) / 3 '(Poly(pie).p(1).Y + Poly(pie).p(2).Y + Poly(pie).p(3).Y) / 3
Else
Ordrd(pie).Y = (Poly(pie).p(1).Y + Poly(pie).p(2).Y + Poly(pie).p(3).Y) / 3
End If
Next pie

QuickSortPolys Ordrd, 1, UBound(Ordrd)
'end sorting

Dim NewRGB As T_RGB
Dim FWD As Boolean
Dim NewLightPos() As T_3DPoint
Dim lightcnt As Integer
Dim LightDistance() As Double
Dim DistancePercent() As Double
Dim LightAnglePct() As Double
Dim LightPercent As T_3DPoint 'x = R, y = G, z = B - cant do 0%-100% in t_rgb

For cyn = UBound(Poly) To 1 Step -1
'this cyn loop goes through every polygon, finds its color, and draws it, then goes to next poly
NewPos = Poly(Ordrd(cyn).X)
'NewPos = Poly(cyn)


NewPos.p(1).X = NewPos.p(1).X - CameraPoint.X
NewPos.p(1).Y = NewPos.p(1).Y - CameraPoint.Y
NewPos.p(1).z = NewPos.p(1).z - CameraPoint.z

NewPos.p(2).X = NewPos.p(2).X - CameraPoint.X
NewPos.p(2).Y = NewPos.p(2).Y - CameraPoint.Y
NewPos.p(2).z = NewPos.p(2).z - CameraPoint.z

NewPos.p(3).X = NewPos.p(3).X - CameraPoint.X
NewPos.p(3).Y = NewPos.p(3).Y - CameraPoint.Y
NewPos.p(3).z = NewPos.p(3).z - CameraPoint.z

NewPos = Rotate3DPoly(NewPos, CameraAngle.X, CameraAngle.Y, CameraAngle.z)
'these  v  6 lines find the 2d position on screen for a 3d triangle.  w0rd.
ScrPos(0).X = (where.ScaleWidth / 2) + ((Angle_GetRadians(NewPos.p(1).X, NewPos.p(1).Y) / CameraWindow.X) * (where.ScaleWidth / 2))
ScrPos(0).Y = (where.ScaleHeight / 2) - ((Angle_GetRadians(NewPos.p(1).z, NewPos.p(1).Y) / CameraWindow.Y) * (where.ScaleHeight / 2))

ScrPos(1).X = (where.ScaleWidth / 2) + ((Angle_GetRadians(NewPos.p(2).X, NewPos.p(2).Y) / CameraWindow.X) * (where.ScaleWidth / 2))
ScrPos(1).Y = (where.ScaleHeight / 2) - ((Angle_GetRadians(NewPos.p(2).z, NewPos.p(2).Y) / CameraWindow.Y) * (where.ScaleHeight / 2))

ScrPos(2).X = (where.ScaleWidth / 2) + ((Angle_GetRadians(NewPos.p(3).X, NewPos.p(3).Y) / CameraWindow.X) * (where.ScaleWidth / 2))
ScrPos(2).Y = (where.ScaleHeight / 2) - ((Angle_GetRadians(NewPos.p(3).z, NewPos.p(3).Y) / CameraWindow.Y) * (where.ScaleHeight / 2))

'below finds the color

'If ((ScrPos(1).Y - ScrPos(0).Y)) > 1 Or ((ScrPos(1).Y - ScrPos(0).Y)) < -1 Then
If NewPos.lighted = True Then

'next nest of if's find which side is facing the camera, so to know if light is visible on that side
If ((Int(ScrPos(1).Y) - Int(ScrPos(0).Y))) <> 0 Then
Slope = ((ScrPos(1).X - ScrPos(0).X) / (ScrPos(1).Y - ScrPos(0).Y))
If Slope > 0 Then
 If ScrPos(2).X - ScrPos(0).X > (((ScrPos(1).X - ScrPos(0).X) / (ScrPos(1).Y - ScrPos(0).Y)) * (ScrPos(2).Y - ScrPos(0).Y)) Then
  If ScrPos(1).Y > ScrPos(0).Y Then
  FWD = True
  Else
  FWD = False
  End If
 Else
  If ScrPos(1).Y < ScrPos(0).Y Then
  FWD = True
  Else
  FWD = False
  End If
 End If
ElseIf Slope < 0 Then
 If ScrPos(2).X - ScrPos(0).X > (((ScrPos(1).X - ScrPos(0).X) / (ScrPos(1).Y - ScrPos(0).Y)) * (ScrPos(2).Y - ScrPos(0).Y)) Then
  If ScrPos(1).Y > ScrPos(0).Y Then
  FWD = True
  Else
  FWD = False
  End If
 Else
  If ScrPos(1).Y < ScrPos(0).Y Then
  FWD = True
  Else
  FWD = False
  End If
 End If
ElseIf Slope = 0 Then
'MsgBox "Slope is " & Slope
End If

ElseIf Int(ScrPos(1).Y) = Int(ScrPos(0).Y) Then
If ScrPos(2).Y <= ScrPos(1).Y Then
 If ScrPos(1).X > ScrPos(0).X Then
 FWD = True
 Else
 FWD = False
 End If
Else
 If ScrPos(1).X < ScrPos(0).X Then
 FWD = True
 Else
 FWD = False
 End If
End If
End If
'sorry if my grammar is not all that, i'm a programmer  ;]  no, i am!
ReDim NewLightPos(UBound(Light))
ReDim LightDistance(UBound(Light))
ReDim LightAnglePct(UBound(Light))
ReDim DistancePercent(UBound(Light))

For lightcnt = 1 To UBound(Light)
polylightdist.X = (((NewPos.p(1).X + CameraPoint.X) + (NewPos.p(2).X + CameraPoint.X) + (NewPos.p(3).X + CameraPoint.X)) / 3) - Light(lightcnt).p.X 'average x coord of poly
polylightdist.Y = (((NewPos.p(1).Y + CameraPoint.Y) + (NewPos.p(2).Y + CameraPoint.Y) + (NewPos.p(3).Y + CameraPoint.Y)) / 3) - Light(lightcnt).p.Y ''' y coord
polylightdist.z = (((NewPos.p(1).z + CameraPoint.z) + (NewPos.p(2).z + CameraPoint.z) + (NewPos.p(3).z + CameraPoint.z)) / 3) - Light(lightcnt).p.z ''' z coord
LightDistance(lightcnt) = Sqr((Sqr((polylightdist.X ^ 2) + (polylightdist.Y ^ 2)) ^ 2) + (polylightdist.z ^ 2))
Next lightcnt

'BZAA!!  Move light according to cameraview!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!  'nm, note to self i guess...
If FWD = True Then
'true angle of light

For lightcnt = 1 To UBound(Light)
If LightDistance(lightcnt) <= Light(lightcnt).dist Then

NewLightPos(lightcnt).X = Light(lightcnt).p.X - CameraPoint.X
NewLightPos(lightcnt).Y = Light(lightcnt).p.Y - CameraPoint.Y
NewLightPos(lightcnt).z = Light(lightcnt).p.z - CameraPoint.z
NewLightPos(lightcnt).z = NewLightPos(lightcnt).z - NewPos.p(1).z
NewLightPos(lightcnt).X = NewLightPos(lightcnt).X - NewPos.p(1).X
NewLightPos(lightcnt).Y = NewLightPos(lightcnt).Y - NewPos.p(1).Y
NewLightPos(lightcnt) = Angle_Rotate3D(NewLightPos(lightcnt).X, NewLightPos(lightcnt).Y, NewLightPos(lightcnt).z, CameraAngle.X, CameraAngle.Y, CameraAngle.z)
End If
Next lightcnt
'find light positions according to poly
NewPos.p(2).z = NewPos.p(2).z - NewPos.p(1).z
NewPos.p(3).z = NewPos.p(3).z - NewPos.p(1).z
NewPos.p(1).z = NewPos.p(1).z - NewPos.p(1).z

NewPos.p(2).X = NewPos.p(2).X - NewPos.p(1).X
NewPos.p(3).X = NewPos.p(3).X - NewPos.p(1).X
NewPos.p(1).X = NewPos.p(1).X - NewPos.p(1).X

NewPos.p(2).Y = NewPos.p(2).Y - NewPos.p(1).Y
NewPos.p(3).Y = NewPos.p(3).Y - NewPos.p(1).Y
NewPos.p(1).Y = NewPos.p(1).Y - NewPos.p(1).Y

RotAngle.X = Angle_GetRadians(NewPos.p(2).Y, NewPos.p(2).z)
RotAngle.X = (pi / 2) - RotAngle.X
NewPos = Rotate3DPoly(NewPos, RotAngle.X, 0, 0)

RotAngle.z = Angle_GetRadians(NewPos.p(2).X, NewPos.p(2).Y)
RotAngle.z = -RotAngle.z
NewPos = Rotate3DPoly(NewPos, 0, 0, RotAngle.z)

RotAngle.Y = Angle_GetRadians(NewPos.p(3).X, NewPos.p(3).z)
RotAngle.Y = (-pi / 2) - RotAngle.Y
NewPos = Rotate3DPoly(NewPos, 0, RotAngle.Y, 0)
'below, rotate lights with poly, so poly is "flat on the ground", and can now calculate angles of light
For lightcnt = 1 To UBound(NewLightPos)
If LightDistance(lightcnt) <= Light(lightcnt).dist Then
NewLightPos(lightcnt) = Angle_Rotate3D(NewLightPos(lightcnt).X, NewLightPos(lightcnt).Y, NewLightPos(lightcnt).z, RotAngle.X, 0, 0)
NewLightPos(lightcnt) = Angle_Rotate3D(NewLightPos(lightcnt).X, NewLightPos(lightcnt).Y, NewLightPos(lightcnt).z, 0, 0, RotAngle.z)
NewLightPos(lightcnt) = Angle_Rotate3D(NewLightPos(lightcnt).X, NewLightPos(lightcnt).Y, NewLightPos(lightcnt).z, 0, RotAngle.Y, 0)
End If
Next lightcnt
'below, calculate angles of light
For lightcnt = 1 To UBound(NewLightPos)
If LightDistance(lightcnt) <= Light(lightcnt).dist Then
LightAnglePct(lightcnt) = 1 - Abs(Angle_GetRadians(Sqr((NewLightPos(lightcnt).X ^ 2) + (NewLightPos(lightcnt).Y ^ 2)), NewLightPos(lightcnt).z) / (pi / 2))
If LightAnglePct(lightcnt) < 0 Then LightAnglePct(lightcnt) = 0
End If
Next lightcnt

NewRGB.R = NewPos.colrmin.R
NewRGB.G = NewPos.colrmin.G
NewRGB.b = NewPos.colrmin.b
LightPercent.X = 0
LightPercent.Y = 0
LightPercent.z = 0
'below, find distance percent, and multiply with angle percent
For lightcnt = 1 To UBound(NewLightPos)
If LightDistance(lightcnt) <= Light(lightcnt).dist Then
DistancePercent(lightcnt) = 1 / ((LightDistance(lightcnt) / Light(lightcnt).halflife) ^ 2)
LightPercent.X = LightPercent.X + (LightAnglePct(lightcnt) * DistancePercent(lightcnt) * Light(lightcnt).val * (Light(lightcnt).colr.R / 255))
LightPercent.Y = LightPercent.Y + (LightAnglePct(lightcnt) * DistancePercent(lightcnt) * Light(lightcnt).val * (Light(lightcnt).colr.G / 255))
LightPercent.z = LightPercent.z + (LightAnglePct(lightcnt) * DistancePercent(lightcnt) * Light(lightcnt).val * (Light(lightcnt).colr.b / 255))
'LightPercent = LightAnglePct(lightcnt) * DistancePercent(lightcnt)
'LightPercent = LightPercent * Light(lightcnt).val
'If LightPercent > 1 Then LightPercent = 1

End If
Next lightcnt
'below, use percentage within colrmin,colr,colrmax to find color
If NewPos.colrmin.R + ((NewPos.colr.R - NewPos.colrmin.R) * LightPercent.X) > NewPos.colr.R Then
NewRGB.R = NewPos.colr.R + ((1 - (1 / (2 ^ (LightPercent.X - 1)))) * (NewPos.colrmax.R - NewPos.colr.R))
Else
NewRGB.R = NewPos.colrmin.R + ((NewPos.colr.R - NewPos.colrmin.R) * LightPercent.X)
End If

If NewPos.colrmin.G + ((NewPos.colr.G - NewPos.colrmin.G) * LightPercent.Y) > NewPos.colr.G Then
NewRGB.G = NewPos.colr.G + ((1 - (1 / (2 ^ (LightPercent.Y - 1)))) * (NewPos.colrmax.G - NewPos.colr.G))
Else
NewRGB.G = NewPos.colrmin.G + ((NewPos.colr.G - NewPos.colrmin.G) * LightPercent.Y)
End If

If NewPos.colrmin.b + ((NewPos.colr.b - NewPos.colrmin.b) * LightPercent.z) > NewPos.colr.b Then
NewRGB.b = NewPos.colr.b + ((1 - (1 / (2 ^ (LightPercent.z - 1)))) * (NewPos.colrmax.b - NewPos.colr.b))
Else
NewRGB.b = NewPos.colrmin.b + ((NewPos.colr.b - NewPos.colrmin.b) * LightPercent.z)
End If


'If NewRGB.R + ((NewPos.colr.R - NewPos.colrmin.R) * LightPercent) > NewPos.colrmax.R Then
'NewRGB.R = NewRGB.R + ((1 - (1 / (2 ^ (((NewPos.colr.R - NewPos.colrmin.R) * LightPercent) / (NewPos.colrmax.R - NewRGB.R))))) * (NewPos.colrmax.R - NewRGB.R))
'Else
'NewRGB.R = NewRGB.R + ((NewPos.colr.R - NewPos.colrmin.R) * LightPercent)
'End If

'If NewRGB.G + ((NewPos.colr.G - NewPos.colrmin.G) * LightPercent) > NewPos.colrmax.G Then
'NewRGB.G = NewRGB.G + ((1 - (1 / (2 ^ (((NewPos.colr.G - NewPos.colrmin.G) * LightPercent) / (NewPos.colrmax.G - NewRGB.G))))) * (NewPos.colrmax.G - NewRGB.G))
'Else
'NewRGB.G = NewRGB.G + ((NewPos.colr.G - NewPos.colrmin.G) * LightPercent)
'End If

'If NewRGB.b + ((NewPos.colr.b - NewPos.colrmin.b) * LightPercent) > NewPos.colrmax.b Then
'NewRGB.b = NewRGB.b + ((1 - (1 / (2 ^ (((NewPos.colr.b - NewPos.colrmin.b) * LightPercent) / (NewPos.colrmax.b - NewRGB.b))))) * (NewPos.colrmax.b - NewRGB.b))
'Else
'NewRGB.b = NewRGB.b + ((NewPos.colr.b - NewPos.colrmin.b) * LightPercent)
'End If



'LightPercent = Abs(Angle_GetRadians(Sqr((NewLightPos(1).X ^ 2) + (NewLightPos(1).Y ^ 2)), NewLightPos(1).z) / (pi / 2))
'LightPercent = 1 - LightPercent

'If LightPercent < 0 Then
'LightPercent = 0
'End If
'NewRGB '''''''
'NewPos.colr.R = 255 * LightPercent
'NewPos.colr.G = 255 * LightPercent
'NewPos.colr.b = 255 * LightPercent


'guess i didnt need all this...  yeah, its been a few weeks since i finished this, so im a little foggy.


'find angle and distance of light



'NewPos.colr.R = 255
'NewPos.colr.G = 255
'NewPos.colr.b = 255

Else
'reverse angle of light  [seeing backside of poly]
ReDim NewLightPos(UBound(Light))

For lightcnt = 1 To UBound(Light)
If LightDistance(lightcnt) <= Light(lightcnt).dist Then

NewLightPos(lightcnt).X = Light(lightcnt).p.X - CameraPoint.X
NewLightPos(lightcnt).Y = Light(lightcnt).p.Y - CameraPoint.Y
NewLightPos(lightcnt).z = Light(lightcnt).p.z - CameraPoint.z
NewLightPos(lightcnt).z = NewLightPos(lightcnt).z - NewPos.p(2).z
NewLightPos(lightcnt).X = NewLightPos(lightcnt).X - NewPos.p(2).X
NewLightPos(lightcnt).Y = NewLightPos(lightcnt).Y - NewPos.p(2).Y
NewLightPos(lightcnt) = Angle_Rotate3D(NewLightPos(lightcnt).X, NewLightPos(lightcnt).Y, NewLightPos(lightcnt).z, CameraAngle.X, CameraAngle.Y, CameraAngle.z)

End If
Next lightcnt


NewPos.p(1).z = NewPos.p(1).z - NewPos.p(2).z
NewPos.p(3).z = NewPos.p(3).z - NewPos.p(2).z
NewPos.p(2).z = NewPos.p(2).z - NewPos.p(2).z

NewPos.p(1).X = NewPos.p(1).X - NewPos.p(2).X
NewPos.p(3).X = NewPos.p(3).X - NewPos.p(2).X
NewPos.p(2).X = NewPos.p(2).X - NewPos.p(2).X

NewPos.p(1).Y = NewPos.p(1).Y - NewPos.p(2).Y
NewPos.p(3).Y = NewPos.p(3).Y - NewPos.p(2).Y
NewPos.p(2).Y = NewPos.p(2).Y - NewPos.p(2).Y


RotAngle.X = Angle_GetRadians(NewPos.p(1).Y, NewPos.p(1).z)
RotAngle.X = (pi / 2) - RotAngle.X
NewPos = Rotate3DPoly(NewPos, RotAngle.X, 0, 0)

RotAngle.z = Angle_GetRadians(NewPos.p(1).X, NewPos.p(1).Y)
RotAngle.z = -RotAngle.z
NewPos = Rotate3DPoly(NewPos, 0, 0, RotAngle.z)

RotAngle.Y = Angle_GetRadians(NewPos.p(3).X, NewPos.p(3).z)
RotAngle.Y = (-pi / 2) - RotAngle.Y
NewPos = Rotate3DPoly(NewPos, 0, RotAngle.Y, 0)

For lightcnt = 1 To UBound(NewLightPos)
If LightDistance(lightcnt) <= Light(lightcnt).dist Then
NewLightPos(lightcnt) = Angle_Rotate3D(NewLightPos(lightcnt).X, NewLightPos(lightcnt).Y, NewLightPos(lightcnt).z, RotAngle.X, 0, 0)
NewLightPos(lightcnt) = Angle_Rotate3D(NewLightPos(lightcnt).X, NewLightPos(lightcnt).Y, NewLightPos(lightcnt).z, 0, 0, RotAngle.z)
NewLightPos(lightcnt) = Angle_Rotate3D(NewLightPos(lightcnt).X, NewLightPos(lightcnt).Y, NewLightPos(lightcnt).z, 0, RotAngle.Y, 0)
End If
Next lightcnt



For lightcnt = 1 To UBound(NewLightPos)
If LightDistance(lightcnt) <= Light(lightcnt).dist Then
LightAnglePct(lightcnt) = 1 - Abs(Angle_GetRadians(Sqr((NewLightPos(lightcnt).X ^ 2) + (NewLightPos(lightcnt).Y ^ 2)), NewLightPos(lightcnt).z) / (pi / 2))
If LightAnglePct(lightcnt) < 0 Then LightAnglePct(lightcnt) = 0
End If
Next lightcnt

NewRGB.R = NewPos.colrmin.R
NewRGB.G = NewPos.colrmin.G
NewRGB.b = NewPos.colrmin.b
LightPercent.X = 0
LightPercent.Y = 0
LightPercent.z = 0

For lightcnt = 1 To UBound(NewLightPos)
If LightDistance(lightcnt) <= Light(lightcnt).dist Then
DistancePercent(lightcnt) = 1 / ((LightDistance(lightcnt) / Light(lightcnt).halflife) ^ 2)
LightPercent.X = LightPercent.X + (LightAnglePct(lightcnt) * DistancePercent(lightcnt) * Light(lightcnt).val * (Light(lightcnt).colr.R / 255))
LightPercent.Y = LightPercent.Y + (LightAnglePct(lightcnt) * DistancePercent(lightcnt) * Light(lightcnt).val * (Light(lightcnt).colr.G / 255))
LightPercent.z = LightPercent.z + (LightAnglePct(lightcnt) * DistancePercent(lightcnt) * Light(lightcnt).val * (Light(lightcnt).colr.b / 255))

End If
Next lightcnt
'If LightPercent > 1 Then LightPercent = 1
If NewPos.colrmin.R + ((NewPos.colr.R - NewPos.colrmin.R) * LightPercent.X) > NewPos.colr.R Then
NewRGB.R = NewPos.colr.R + ((1 - (1 / (2 ^ (LightPercent.X - 1)))) * (NewPos.colrmax.R - NewPos.colr.R))
Else
NewRGB.R = NewPos.colrmin.R + ((NewPos.colr.R - NewPos.colrmin.R) * LightPercent.X)
End If

If NewPos.colrmin.G + ((NewPos.colr.G - NewPos.colrmin.G) * LightPercent.Y) > NewPos.colr.G Then
NewRGB.G = NewPos.colr.G + ((1 - (1 / (2 ^ (LightPercent.Y - 1)))) * (NewPos.colrmax.G - NewPos.colr.G))
Else
NewRGB.G = NewPos.colrmin.G + ((NewPos.colr.G - NewPos.colrmin.G) * LightPercent.Y)
End If

If NewPos.colrmin.b + ((NewPos.colr.b - NewPos.colrmin.b) * LightPercent.z) > NewPos.colr.b Then
NewRGB.b = NewPos.colr.b + ((1 - (1 / (2 ^ (LightPercent.z - 1)))) * (NewPos.colrmax.b - NewPos.colr.b))
Else
NewRGB.b = NewPos.colrmin.b + ((NewPos.colr.b - NewPos.colrmin.b) * LightPercent.z)
End If

'If NewRGB.R + ((NewPos.colr.R - NewPos.colrmin.R) * LightPercent) > NewPos.colrmax.R Then
'NewRGB.R = NewRGB.R + ((1 - (1 / (2 ^ (((NewPos.colr.R - NewPos.colrmin.R) * LightPercent) / (NewPos.colrmax.R - NewRGB.R))))) * (NewPos.colrmax.R - NewRGB.R))
'Else
'NewRGB.R = NewRGB.R + ((NewPos.colr.R - NewPos.colrmin.R) * LightPercent)
'End If

'If NewRGB.G + ((NewPos.colr.G - NewPos.colrmin.G) * LightPercent) > NewPos.colrmax.G Then
'NewRGB.G = NewRGB.G + ((1 - (1 / (2 ^ (((NewPos.colr.G - NewPos.colrmin.G) * LightPercent) / (NewPos.colrmax.G - NewRGB.G))))) * (NewPos.colrmax.G - NewRGB.G))
'Else
'NewRGB.G = NewRGB.G + ((NewPos.colr.G - NewPos.colrmin.G) * LightPercent)
'End If

'If NewRGB.b + ((NewPos.colr.b - NewPos.colrmin.b) * LightPercent) > NewPos.colrmax.b Then
'NewRGB.b = NewRGB.b + ((1 - (1 / (2 ^ (((NewPos.colr.b - NewPos.colrmin.b) * LightPercent) / (NewPos.colrmax.b - NewRGB.b))))) * (NewPos.colrmax.b - NewRGB.b))
'Else
'NewRGB.b = NewRGB.b + ((NewPos.colr.b - NewPos.colrmin.b) * LightPercent)
'End If



End If
'AD Int(LightPercent * 100)

'AD NewRGB.b
'Pause 0.1
'If NewRGB.R > 255 Then NewRGB.R = 255
'If NewRGB.G > 255 Then NewRGB.G = 255
'If NewRGB.b > 255 Then NewRGB.b = 255

NewPos.colr.R = NewRGB.R
NewPos.colr.G = NewRGB.G
NewPos.colr.b = NewRGB.b


End If  'NewPos.lighted = True
'stop finding color!
'AD LightAnglePct(1)
'AD Int(LightPercent * 100)
'If FWD = True Then
'AD NewPos.colr.b
'Pause 0.1
'AD NewPos.colr.b

'yeah, go ahead, use drawtriangle instead of drawtri3, see if i care!
'DrawTriangle where, ScrPos, RGB(NewPos.colr.R, NewPos.colr.G, NewPos.colr.b)

'the almighty drawing of thine colored triangle!
drawtri3 where, ScrPos, RGB(NewPos.colr.R, NewPos.colr.G, NewPos.colr.b), NewPos.outlined
'If Int(Timer) Mod 2 = 1 Then
'If NewPos.outlined = True Then


'DrawTriangle2 where, ScrPos, RGB(OutlinedColor.R, OutlinedColor.G, OutlinedColor.b)
'End If
'DrawTriangle2 where, ScrPos, RGB(0, 0, 255)
'End If
Next

End Sub

Public Sub QuickSortPolys(vArray() As T_Point, inLow As Long, inHi As Long)
'sniff, i didnt make this... thanks whoever came up with the quicksort ;]
   Dim pivot   As Double
   Dim tmpSwap As T_Point
   Dim tmpLow  As Long
   Dim tmpHi   As Long
    
   tmpLow = inLow
   tmpHi = inHi
    
   pivot = vArray((inLow + inHi) \ 2).Y
  
   While (tmpLow <= tmpHi)
  
      While (vArray(tmpLow).Y < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < vArray(tmpHi).Y And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      If (tmpLow <= tmpHi) Then
         tmpSwap.Y = vArray(tmpLow).Y
         tmpSwap.X = vArray(tmpLow).X
         vArray(tmpLow).X = vArray(tmpHi).X
         vArray(tmpLow).Y = vArray(tmpHi).Y
         vArray(tmpHi).X = tmpSwap.X
         vArray(tmpHi).Y = tmpSwap.Y
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then QuickSortPolys vArray, inLow, tmpHi
   If (tmpLow < inHi) Then QuickSortPolys vArray, tmpLow, inHi
  
End Sub
Function Rotate3DPoly(ByRef d As T_3DPoly, xang, yang, zang) As T_3DPoly
'well, for a system of grouped polys, like a game or something, you'll need a system of
'rotation, meaning that there are 0 to infinity levels of rotation per polygon, so the lower-levels
'of rotation are applied first.  The engine will not account for these levels, they must be handled by
'the game code itself...  you will probably need a variable to act as poly(), and the actual poly() will be
'the rotated polys...  and you will probably make a variable [polyrotate().level().x,y,z] to handle each poly's
'leveled rotations... and you'll prolly find a way to manage rotations using groups of polys, since an object is
'many polys.... w00t   oh, this function rotates 3 points ;]
Rotate3DPoly.colr.R = d.colr.R
Rotate3DPoly.colr.G = d.colr.G
Rotate3DPoly.colr.b = d.colr.b
Rotate3DPoly.lighted = d.lighted
Rotate3DPoly.colrmin.R = d.colrmin.R
Rotate3DPoly.colrmin.G = d.colrmin.G
Rotate3DPoly.colrmin.b = d.colrmin.b
Rotate3DPoly.colrmax.R = d.colrmax.R
Rotate3DPoly.colrmax.G = d.colrmax.G
Rotate3DPoly.colrmax.b = d.colrmax.b
Rotate3DPoly.outlined = d.outlined



Dim length, newrads
length = Sqr((d.p(1).Y ^ 2) + (d.p(1).z ^ 2))
newrads = Angle_GetRadians(d.p(1).Y, d.p(1).z) + xang
Rotate3DPoly.p(1).Y = Sin(newrads) * length
Rotate3DPoly.p(1).z = Cos(newrads) * length
length = Sqr((d.p(1).X ^ 2) + (Rotate3DPoly.p(1).z ^ 2))
newrads = Angle_GetRadians(d.p(1).X, Rotate3DPoly.p(1).z) + yang
Rotate3DPoly.p(1).X = Sin(newrads) * length
Rotate3DPoly.p(1).z = Cos(newrads) * length
length = Sqr((Rotate3DPoly.p(1).X ^ 2) + (Rotate3DPoly.p(1).Y ^ 2))
newrads = Angle_GetRadians(Rotate3DPoly.p(1).X, Rotate3DPoly.p(1).Y) + zang
Rotate3DPoly.p(1).X = Sin(newrads) * length
Rotate3DPoly.p(1).Y = Cos(newrads) * length

length = Sqr((d.p(2).Y ^ 2) + (d.p(2).z ^ 2))
newrads = Angle_GetRadians(d.p(2).Y, d.p(2).z) + xang
Rotate3DPoly.p(2).Y = Sin(newrads) * length
Rotate3DPoly.p(2).z = Cos(newrads) * length
length = Sqr((d.p(2).X ^ 2) + (Rotate3DPoly.p(2).z ^ 2))
newrads = Angle_GetRadians(d.p(2).X, Rotate3DPoly.p(2).z) + yang
Rotate3DPoly.p(2).X = Sin(newrads) * length
Rotate3DPoly.p(2).z = Cos(newrads) * length
length = Sqr((Rotate3DPoly.p(2).X ^ 2) + (Rotate3DPoly.p(2).Y ^ 2))
newrads = Angle_GetRadians(Rotate3DPoly.p(2).X, Rotate3DPoly.p(2).Y) + zang
Rotate3DPoly.p(2).X = Sin(newrads) * length
Rotate3DPoly.p(2).Y = Cos(newrads) * length

length = Sqr((d.p(3).Y ^ 2) + (d.p(3).z ^ 2))
newrads = Angle_GetRadians(d.p(3).Y, d.p(3).z) + xang
Rotate3DPoly.p(3).Y = Sin(newrads) * length
Rotate3DPoly.p(3).z = Cos(newrads) * length
length = Sqr((d.p(3).X ^ 2) + (Rotate3DPoly.p(3).z ^ 2))
newrads = Angle_GetRadians(d.p(3).X, Rotate3DPoly.p(3).z) + yang
Rotate3DPoly.p(3).X = Sin(newrads) * length
Rotate3DPoly.p(3).z = Cos(newrads) * length
length = Sqr((Rotate3DPoly.p(3).X ^ 2) + (Rotate3DPoly.p(3).Y ^ 2))
newrads = Angle_GetRadians(Rotate3DPoly.p(3).X, Rotate3DPoly.p(3).Y) + zang
Rotate3DPoly.p(3).X = Sin(newrads) * length
Rotate3DPoly.p(3).Y = Cos(newrads) * length



End Function


