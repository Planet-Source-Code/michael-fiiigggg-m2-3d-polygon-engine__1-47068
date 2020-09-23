Attribute VB_Name = "MgFmod"
Public holdb()
'well, sorry if my code isnt very "clean", and not
'superbly commented.  I hadn't planned on releasing
'this, but once I made the 3d engine, I figured
'this might help someone out  ;]  Well, there is
'frmpoly, 3dfunctions, and mgfmod..   mgfmod is
'just a collection of functions i've made throughout
'programming, most not relating to 3d.  I do have some
'functions here used by this 3d program though.  in 3dfunctions,
'there is most of the 3d engine for drawing everything.
'the form is just an instance of what can be done using the
'engine.  All relevant parts of the engine, I will comment, and
'in each small function, i'll label what that function does.
'you can use this engine if you want...  I'm sure there are much
'better engines in c++ / directX, which is what I plan to get into
'next.  If you do use this engine for a project, please give
'credit where credit is due.  I did all the code here, except for
'the quicksort of polygons - that I needed the most efficient way, not
'my own jumbled code =p.  enjoy.
'made by michael figamarellay [name shizzled to protect the innocent]
'IM me on aim, m2g2f8 if ya have any questions


Option Compare Text
Option Explicit
Dim yes
Public Const pi = 3.1415926535
'for mgf3d v1.0
Public xang As Double
Public yang As Double
Public zang As Double
Public LinePos() As T_3DLineCoordinates
Public ZoomFact As Double
Public RotateAllAngle As T_3DPoint
Public ViewPoint As T_3DPoint
Public ViewAngle As T_3DPoint
Public AutoViewLook As Boolean
Public ViewWindow As T_3DPoint
Type T_3DLineCoordinates
 x1 As Variant
 y1 As Variant
 z1 As Variant
 x2 As Variant
 y2 As Variant
 z2 As Variant
End Type
Type T_3DPoint
 X As Variant
 Y As Variant
 z As Variant
End Type
Type T_3DViewWindow
 X As Variant
 z As Variant
End Type

Type T_LineCoordinates
 x1 As Variant
 y1 As Variant
 x2 As Variant
 y2 As Variant
End Type
Type T_bla
 lines() As T_LineCoordinates
End Type
Type T_Point
 X As Variant
 Y As Variant
End Type
Type T_RGB
 R As Long
 G As Long
 b As Long
End Type

Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Global Const IcD_ADD = &H0&
Global Const IcD_MODIFY = &H1
Global Const IcD_DELETE = &H2
Global Const IcD_BACKMESSAGE = &H1
Global Const IcD_ICON = &H2
Global Const IcD_TIP = &H4
Global Const WM_MOUSEMOVE = &H200
Global IcD As NOTIFYICONDATA
Const flags = SWP_NOMOVE Or SWP_NOSIZE
Type POINTAPI
 X As Long
 Y As Long
End Type
Type NOTIFYICONDATA
 cbze As Long
 hwnd As Long
 exid As Long
 flags As Long
 backmessage As Long
 icon As Long
 tip As String * 64
End Type
Type T_GridStuff
On As Boolean
Num As String
End Type

Type T_GridY
Y() As T_GridStuff
End Type

Type T_Grid
X() As T_GridY
Startc As T_Point
Destc As T_Point
End Type





Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const DSINVERT = &H550009
Public Const NOTSRCCOPY = &H330008
Public Const NOTSRCERASE = &H1100A6
Public Const MERGECOPY = &HC000CA
Public Const MERGEPAINT = &HBB0226
Public Const PATCOPY = &HF00021
Public Const PATINVERT = &H5A0049
Public Const PATPAINT = &HFB0A09
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42



Function Angle_GetRadians(X, Y)
'since atn(x), sometimes x is something divided by zero, cant have that  ;]
Dim ATNnum
If RemEx(Y) = 0 Then
If X = 0 Then
ATNnum = 0
Else
If X > 0 Then
ATNnum = pi / 2
Else
ATNnum = -pi / 2
End If
End If
Else
ATNnum = RemEx(Atn(RemEx(X) / RemEx(Y)))
End If
If X >= 0 And Y >= 0 Then Angle_GetRadians = ATNnum
If X >= 0 And Y < 0 Then Angle_GetRadians = ATNnum + pi
If X < 0 And Y < 0 Then Angle_GetRadians = ATNnum - pi
If X < 0 And Y >= 0 Then Angle_GetRadians = ATNnum
End Function
Sub Create3DLine(x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double, Color As Double)
'used in old 3dline program
ReDim Preserve LinePos(UBound(LinePos) + 1)
LinePos(UBound(LinePos)).x1 = x1
LinePos(UBound(LinePos)).y1 = y1
LinePos(UBound(LinePos)).z1 = z1
LinePos(UBound(LinePos)).x2 = x2
LinePos(UBound(LinePos)).y2 = y2
LinePos(UBound(LinePos)).z2 = z2
'Load Form1.LineShow(UBound(LinePos))
'UpdateDisplay
'Form1.LineShow(UBound(LinePos)).BorderColor = Color
'Form1.LineShow(UBound(LinePos)).Visible = True
End Sub

Sub DistSizeFilter(ByRef Xpos, ByRef Ypos, ByRef Zpos, DSMval)
'uhm, i forgot.  figure it out, if you care to =p
If Ypos >= 0 Then
Xpos = (1 - (Ypos / DSMval)) * Xpos
Zpos = (1 - (Ypos / DSMval)) * Zpos
If Ypos > DSMval Then
Xpos = 0: Zpos = 0
End If
Else
Xpos = (1 + (Ypos / -DSMval)) * Xpos
Zpos = (1 + (Ypos / -DSMval)) * Zpos
End If
End Sub

Sub DrawTriangle(where As Object, Coords() As T_Point, colr As Long)
'my original sub to drawtriangle... a little buggy - but not used now,
'since i was told about the drawpolygon api, which is faster and looks better


'coords = 0, 1, and 2
'AD "c=" & Coords(0).x
'AD "c=" & Coords(0).y
'AD "c=" & Coords(1).x
'AD "c=" & Coords(1).y
'AD "c=" & Coords(2).x
'AD "c=" & Coords(2).y


Dim L As Integer
Dim Ordered(3) As Integer
Dim NCoords(3) As T_Point
'order the coords
If (Coords(0).Y <= Coords(1).Y And Coords(0).Y <= Coords(2).Y) Then
Ordered(0) = 0
ElseIf (Coords(1).Y <= Coords(0).Y And Coords(1).Y <= Coords(2).Y) Then
Ordered(0) = 1
ElseIf (Coords(2).Y <= Coords(0).Y And Coords(2).Y <= Coords(1).Y) Then
Ordered(0) = 2
End If

If ((Coords(0).Y > Coords(1).Y And Coords(0).Y < Coords(2).Y) Or (Coords(0).Y < Coords(1).Y And Coords(0).Y > Coords(2).Y)) Then
Ordered(1) = 0
ElseIf ((Coords(1).Y > Coords(0).Y And Coords(1).Y < Coords(2).Y) Or (Coords(1).Y < Coords(0).Y And Coords(1).Y > Coords(2).Y)) Then
Ordered(1) = 1
ElseIf ((Coords(2).Y > Coords(1).Y And Coords(2).Y < Coords(0).Y) Or (Coords(2).Y < Coords(1).Y And Coords(2).Y > Coords(0).Y)) Then
Ordered(1) = 2
End If

If (Coords(0).Y >= Coords(1).Y And Coords(0).Y >= Coords(2).Y) Then
Ordered(2) = 0
ElseIf (Coords(1).Y >= Coords(0).Y And Coords(1).Y >= Coords(2).Y) Then
Ordered(2) = 1
ElseIf (Coords(2).Y >= Coords(0).Y And Coords(2).Y >= Coords(1).Y) Then
Ordered(2) = 2
End If

'fix the stupid error since i suck at making 100% correct ifs  =[

Dim p35, p66
Dim NuE(3) As Boolean
Dim NuN(3) As Integer
NuE(0) = False
NuE(1) = False
NuE(2) = False
For p35 = 0 To 2
If Ordered(p35) = 0 Then
NuE(0) = True
NuN(0) = NuN(0) + 1
ElseIf Ordered(p35) = 1 Then
NuE(1) = True
NuN(1) = NuN(1) + 1
ElseIf Ordered(p35) = 2 Then
NuE(2) = True
NuN(2) = NuN(2) + 1
End If
Next p35


If NuE(1) = False Then
If NuN(0) > 1 Then
For p66 = 0 To 2
If Ordered(p66) = 0 Then
Ordered(p66) = 1
Exit For
End If
Next p66
End If
ElseIf NuE(2) = False Then
If NuN(0) > 1 Then
For p66 = 0 To 2
If Ordered(p66) = 0 Then
Ordered(p66) = 2
Exit For
End If
Next p66
End If
End If
'001, 002, 020, 200, 100, 010

'done fixed

 

NCoords(0).X = Coords(Ordered(0)).X
NCoords(0).Y = Coords(Ordered(0)).Y
NCoords(1).X = Coords(Ordered(1)).X
NCoords(1).Y = Coords(Ordered(1)).Y
NCoords(2).X = Coords(Ordered(2)).X
NCoords(2).Y = Coords(Ordered(2)).Y
'done ordering

Dim NPStep As Integer
If (NCoords(1).Y - NCoords(0).Y) > 1 And (NCoords(2).Y - NCoords(0).Y) > 1 Then
For L = 0 To (NCoords(1).Y - NCoords(0).Y)

where.Line (NCoords(0).X + ((NCoords(1).X - NCoords(0).X) * (L / (NCoords(1).Y - NCoords(0).Y))), NCoords(0).Y + L)- _
(NCoords(0).X + ((NCoords(2).X - NCoords(0).X) * (L / (NCoords(2).Y - NCoords(0).Y))), NCoords(0).Y + L), colr
Next L
End If
'If (NCoords(2).y - NCoords(1).y) <> 0 Then
'MsgBox NCoords(2).y - NCoords(1).y
If (NCoords(2).Y - NCoords(1).Y) > 1 And (NCoords(2).Y - NCoords(0).Y) > 1 Then
For L = 0 To (NCoords(2).Y - NCoords(1).Y)
where.Line (NCoords(1).X + ((NCoords(2).X - NCoords(1).X) * (L / (NCoords(2).Y - NCoords(1).Y))), NCoords(1).Y + L)- _
(NCoords(0).X + ((NCoords(2).X - NCoords(0).X) * ((NCoords(1).Y - NCoords(0).Y + L) / (NCoords(2).Y - NCoords(0).Y))), NCoords(1).Y + L), colr
Next L
End If





End Sub

Sub File_Compress(filename As String, compress As Boolean)
'erhm, i dont think this fully works...  not algorythmic compressor, just an unused-char compressor
Dim lc, lc2, found As Boolean, base As Integer, base2 As Integer, thebin As String
Dim hold() As Byte, hold2() As Byte, hold3() As String, hold4() As String, hold5() As Byte, hold6() As Byte, wha As Integer
Dim finaltxt As String
If compress = True Then
Open filename For Binary Access Read As #1
ReDim hold(FileLen(filename))
 For lc = 1 To FileLen(filename)
 If EOF(1) = True Then Exit For
 Get 1, lc, hold(lc)
 Next lc
ReDim hold2(0)
 For lc = 1 To UBound(hold)
 found = False
  For lc2 = 1 To UBound(hold2)
   If hold(lc) = hold2(lc2) Then
   found = True
   Exit For
   End If
  Next lc2
  If found = False Then
  ReDim Preserve hold2(UBound(hold2) + 1)
  hold2(UBound(hold2)) = hold(lc)
  End If
 Next lc
base = 0
 Do
 base = base + 1
 If 2 ^ base >= 256 Then
 Exit Do
 End If
 Loop
ReDim hold3(2 ^ base - 1)
 For lc = 1 To base
 hold3(0) = hold3(0) & "0"
 Next lc
 For lc = 1 To 2 ^ base - 1
 hold3(lc) = hold3(lc - 1)
 hold3(lc) = Left(hold3(lc), base - 1) & val(Right(hold3(lc), 1)) + 1
  DoEvents
  For lc2 = base To 1 Step -1
   If val(Mid(hold3(lc), lc2, 1)) > 1 Then
   hold3(lc) = Left$(hold3(lc), lc2 - 2) & (val(Mid(hold3(lc), lc2 - 1, 1)) + 1) & "0" & Mid(hold3(lc), lc2 + 1)
   End If
  Next lc2
 Next lc
ReDim hold4(0)
base2 = 0
 Do
 base2 = base2 + 1
 If 2 ^ base2 >= UBound(hold2) Then
 Exit Do
 End If
 Loop
ReDim hold4(2 ^ base2 - 1)
 For lc = 1 To base2
 hold4(0) = hold4(0) & "0"
 Next lc
 For lc = 1 To 2 ^ base2 - 1
 hold4(lc) = hold4(lc - 1)
 hold4(lc) = Left(hold4(lc), base2 - 1) & val(Right(hold4(lc), 1)) + 1
  DoEvents
  For lc2 = base2 To 1 Step -1
   If val(Mid(hold4(lc), lc2, 1)) > 1 Then
   hold4(lc) = Left$(hold4(lc), lc2 - 2) & (val(Mid(hold4(lc), lc2 - 1, 1)) + 1) & "0" & Mid(hold4(lc), lc2 + 1)
   End If
  Next lc2
 Next lc
'hold = original txt
'hold2 = found txt
'hold3 = 256 binary
'hold4 = <256 binary

ReDim hold6(1)
hold6(1) = UBound(hold2)

For lc = 1 To UBound(hold2)
ReDim Preserve hold6(UBound(hold6) + 1)
hold6(UBound(hold6)) = hold2(lc)
Next lc



For lc = 1 To UBound(hold)
For lc2 = 1 To UBound(hold2)
If hold(lc) = hold2(lc2) Then
thebin = thebin & hold4(lc2 - 1)
Exit For
End If
Next lc2
Next lc

lc2 = Len(thebin)
For lc = 1 To lc2 Mod 8
thebin = thebin & "0"
Next lc
ReDim hold5(0)
For lc = 1 To Len(thebin) Step 8
ReDim Preserve hold5(UBound(hold5) + 1)
For lc2 = 0 To UBound(hold3)
If Mid$(thebin, lc, 8) = hold3(lc2) Then
hold5(UBound(hold5)) = lc2
Exit For
End If
Next lc2
Next lc


For lc = 1 To UBound(hold6)
AD Chr$(hold6(lc)), False
Next lc

For lc = 1 To UBound(hold5)
AD Chr$(hold5(lc)), False
Next lc

ElseIf compress = False Then

Open filename For Binary Access Read As #1
ReDim hold(FileLen(filename))
 For lc = 1 To FileLen(filename)
 If EOF(1) = True Then Exit For
 Get 1, lc, hold(lc)
 Next lc
ReDim hold2(0)
ReDim hold2(hold(1))
For lc = 2 To hold(1) + 1
hold2(lc - 1) = hold(lc)
Next lc
ReDim hold6(0)

For lc = hold(1) + 1 To UBound(hold)
ReDim Preserve hold6(UBound(hold6) + 1)
hold6(UBound(hold6)) = hold(lc)
Next lc



base = 0
 Do
 base = base + 1
 If 2 ^ base >= 256 Then
 Exit Do
 End If
 Loop
ReDim hold4(2 ^ base - 1)
 For lc = 1 To base
 hold4(0) = hold4(0) & "0"
 Next lc
 For lc = 1 To 2 ^ base - 1
 hold4(lc) = hold4(lc - 1)
 hold4(lc) = Left(hold4(lc), base - 1) & val(Right(hold4(lc), 1)) + 1
  DoEvents
  For lc2 = base To 1 Step -1
   If val(Mid(hold4(lc), lc2, 1)) > 1 Then
   hold4(lc) = Left$(hold4(lc), lc2 - 2) & (val(Mid(hold4(lc), lc2 - 1, 1)) + 1) & "0" & Mid(hold4(lc), lc2 + 1)
   End If
  Next lc2
 Next lc


ReDim hold3(0)
base2 = 0
 Do
 base2 = base2 + 1
 If 2 ^ base2 >= UBound(hold2) Then
 Exit Do
 End If
 Loop
ReDim hold3(2 ^ base2 - 1)
 For lc = 1 To base2
 hold3(0) = hold3(0) & "0"
 Next lc
 For lc = 1 To 2 ^ base2 - 1
 hold3(lc) = hold3(lc - 1)
 hold3(lc) = Left(hold3(lc), base2 - 1) & val(Right(hold3(lc), 1)) + 1
  DoEvents
  For lc2 = base2 To 1 Step -1
   If val(Mid(hold3(lc), lc2, 1)) > 1 Then
   hold3(lc) = Left$(hold3(lc), lc2 - 2) & (val(Mid(hold3(lc), lc2 - 1, 1)) + 1) & "0" & Mid(hold3(lc), lc2 + 1)
   End If
  Next lc2
 Next lc
'hold = file
'hold2 = used chars
'hold6 = compressed binary
'hold4 = 256 binary
'hold5 = <256 binary
For lc = 1 To UBound(hold3)
thebin = thebin & hold4(hold6(lc))
Next lc

lc = 0
Do
lc = lc + 1
If 2 ^ lc >= UBound(hold3) + 1 Then
wha = lc
Exit Do
End If
Loop

For lc = 1 To Len(thebin) Step wha
For lc2 = 0 To UBound(hold3)
If Mid$(thebin, lc, wha) = hold3(lc2) Then
finaltxt = finaltxt & Chr$(hold2(lc2 + 1))
Exit For
End If
Next lc2
Next lc
AD finaltxt

End If
Close #1
End Sub

Function Num_Round(Num)
If Num < 0 Then
Num = 0 - Num
Num_Round = -Int(Num + 0.5)
Else
Num_Round = Int(Num + 0.5)
End If
End Function
Sub AD(ByVal txt, Optional enterbefore As Variant)
'form1.text1 - used to debug programs if theres a bug, and i need to check values
If IsMissing(enterbefore) = True Then
Form1.Text1.Text = Form1.Text1.Text & Chr$(13) & Chr$(10) & txt
ElseIf IsMissing(enterbefore) = False And enterbefore = False Then
Form1.Text1.Text = Form1.Text1.Text & txt
ElseIf IsMissing(enterbefore) = False And enterbefore = True Then
Form1.Text1.Text = Form1.Text1.Text & Chr$(13) & Chr$(10) & txt
End If
Form1.Text1.SelStart = Len(Form1.Text1.Text)
End Sub

Sub Array_ClearNulls(ByRef Ary())
'removes all "" in array
Dim lc1 As Long, lc2 As Long, Done As Boolean, Max As Long
Max = UBound(Ary)
lc1 = 0: lc2 = 0: Done = False
Do Until Done = True
lc1 = lc1 + 1
If Ary(lc1) = "" Then
If lc1 >= Max Then
ReDim Preserve Ary(lc1 - 1)
Done = True
Exit Do
End If
If lc1 = Max Then
Max = Max - 1
ReDim Preserve Ary(Max - 1)
Else
For lc2 = lc1 To Max - 1
Ary(lc2) = Ary(lc2 + 1)
Next lc2
Max = Max - 1
ReDim Preserve Ary(Max)
End If
End If
If Max >= lc1 Then
If Ary(lc1) = "" Then lc1 = lc1 - 1
End If
If lc1 >= Max Then Done = True
Loop
End Sub


Sub Array_OpenFile(ByVal filename As String, ByRef Ary())
'opens file to array
Dim FileHold As String, lc1 As Long
lc1 = 0
Open filename For Input As #1
Do Until EOF(1) = True
Line Input #1, FileHold
DoEvents
lc1 = lc1 + 1
ReDim Preserve Ary(lc1)
Ary(lc1) = FileHold
Loop
Close #1
End Sub

Sub Array_SaveFile(ByVal filename As String, Ary())
'saves file from array
Dim lc1 As Long, FileHold
Open filename For Output As #1
For lc1 = 1 To UBound(Ary)
FileHold = FileHold & Ary(lc1) & Chr$(13) & Chr$(10)
DoEvents
Next lc1
Print #1, FileHold
DoEvents
Close #1
End Sub

Sub Form_TrayOff()
'turn off tray icon
Shell_NotifyIconA IcD_DELETE, IcD
End Sub

Sub Form_TrayOn(pic As Object, tooltip As String)
'make a picturebox, put a picture in it (picture1.picture)... then put code in the mousemove event
'put the following commented code into the picturebox's mousemove event.

'Dim tray As Long
'tray = (X And &HFF) * &H100
'Select Case tray
'Case 0 'mouse moves
'Case &HF00 'left mouse button down
'Case &H1E00 'left mouse button up
'Case &H2D00 'left mouse button double click
'Case &H3C00 'right mouse button down
'Case &H4B00 'right mouse button up
'Case &H5A00 'right mouse button double click
'End Select

IcD.cbze = Len(IcD)
IcD.hwnd = pic.hwnd
IcD.exid = 0
IcD.exid = IcD.exid + 1
IcD.flags = IcD_BACKMESSAGE Or IcD_ICON Or IcD_TIP
IcD.icon = pic.Picture
IcD.tip = tooltip & Chr$(0)
IcD.backmessage = WM_MOUSEMOVE
'lResult = Shell_NotifyIconA(IcDM_ADD, IcD)
End Sub

Sub Form_OnTop(frm As Form)
SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags
End Sub
Sub Form_NotOnTop(frm As Form)
'okay, no more obvious comments!
Const wFlags = SWP_NOMOVE Or SWP_NOSIZE
SetWindowPos frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, wFlags
End Sub
Sub DeleteDirectory(DirName As String)
'duh
RmDir DirName
End Sub

Function DirectoryExists(filename) As Boolean
'..
Dim DE As String
DE = Dir((filename), 16)
If DE <> "" Then
DirectoryExists = True
Else
DirectoryExists = False
End If
End Function





Function FileExists(filename) As Boolean
'checks if a file is foshizzle'
Dim FE As String
FE = Dir((filename), 6)
If FE <> "" Then
FileExists = True
Else
FileExists = False
End If
End Function



Sub Form_CenterScreen(whichform As Object)
'yes, send it j00r form!
whichform.Left = (Screen.Width - whichform.Width) / 2
whichform.Top = (Screen.Height - whichform.Height) / 2
End Sub

Sub Form_FadeBackColor(thefrm As Object, Color1, Color2, NumTimes As Long)
'fades, or something...  shit, always comment your functions, kids...
Dim Red1, Red2, Red3
Dim Green1, Green2, Green3
Dim Blue1, Blue2, Blue3
Dim lc1
Dim lc2
Red1 = GetRGB(Color1, "R")
Green1 = GetRGB(Color1, "G")
Blue1 = GetRGB(Color1, "B")
Red2 = GetRGB(Color2, "R") - GetRGB(Color1, "R")
Green2 = GetRGB(Color2, "G") - GetRGB(Color1, "G")
Blue2 = GetRGB(Color2, "B") - GetRGB(Color1, "B")
Do
DoEvents
lc1 = lc1 + 1
Red3 = Int(Red1 + (Red2 * (lc1 / NumTimes)))
Green3 = Int(Green1 + (Green2 * (lc1 / NumTimes)))
Blue3 = Int(Blue1 + (Blue2 * (lc1 / NumTimes)))
thefrm.BackColor = RGB(Red3, Green3, Blue3)
If lc1 >= NumTimes Then Exit Do
Loop
End Sub





Sub Form_FadeColor(thefrm As Object, ByRef AryBR())
'w00ha!'  perhaps this was for my 2d line screensaver
Dim Ary()
Dim DivBy As Long
Dim lines() As T_LineCoordinates
Dim LineD() As T_LineCoordinates
Dim LineSC() As T_RGB
Dim LineDC() As T_RGB
Dim LineType() As String
Dim LineSpeed As Long
Static LastS() As T_LineCoordinates
Static LastD() As T_LineCoordinates
Static LastSC() As T_RGB
Static LastDC() As T_RGB
Static DoneYet As Boolean
Dim lc1, lc2, lc3, LC4
Dim W1, H1
W1 = thefrm.ScaleWidth
H1 = thefrm.ScaleHeight
ReDim Ary(UBound(AryBR))
For lc1 = 1 To UBound(Ary)
Ary(lc1) = AryBR(lc1)
Next lc1
ReDim lines((UBound(Ary) - 2) / 11)
ReDim LineD((UBound(Ary) - 2) / 11)
ReDim LineSC((UBound(Ary) - 2) / 11)
ReDim LineDC((UBound(Ary) - 2) / 11)
ReDim LineType((UBound(Ary) - 2) / 11)
LineSpeed = Ary(1)
DivBy = Ary(2) - 1
If DoneYet = False Then
DoneYet = True
ReDim LastS(UBound(lines))
ReDim LastD(UBound(lines))
ReDim LastSC(UBound(lines))
ReDim LastDC(UBound(lines))
For lc1 = 1 To ((UBound(Ary) - 2) / 11)
LastS(lc1).x1 = RndNum(DivBy) / DivBy
LastS(lc1).y1 = RndNum(DivBy) / DivBy
LastS(lc1).x2 = RndNum(DivBy) / DivBy
LastS(lc1).y2 = RndNum(DivBy) / DivBy
LastD(lc1).x1 = RndNum(DivBy) / DivBy
LastD(lc1).y1 = RndNum(DivBy) / DivBy
LastD(lc1).x2 = RndNum(DivBy) / DivBy
LastD(lc1).y2 = RndNum(DivBy) / DivBy
LastSC(lc1).R = RndNum(255)
LastSC(lc1).G = RndNum(255)
LastSC(lc1).b = RndNum(255)
LastDC(lc1).R = RndNum(255)
LastDC(lc1).G = RndNum(255)
LastDC(lc1).b = RndNum(255)
Next lc1
End If
For lc1 = 1 To ((UBound(Ary) - 2) / 11)
If IsNumeric(Ary(((lc1 - 1) * 11) + 3)) = True Then
LineSC(lc1).R = GetRGB(Ary(((lc1 - 1) * 11) + 3), "R")
LineSC(lc1).G = GetRGB(Ary(((lc1 - 1) * 11) + 3), "G")
LineSC(lc1).b = GetRGB(Ary(((lc1 - 1) * 11) + 3), "B")
Else
Select Case Ary(((lc1 - 1) * 11) + 3)
Case ":Random"
LineSC(lc1).R = RndNum(255)
LineSC(lc1).G = RndNum(255)
LineSC(lc1).b = RndNum(255)
Case ":LastSC"
LineSC(lc1).R = LastSC(lc1).R
LineSC(lc1).G = LastSC(lc1).G
LineSC(lc1).b = LastSC(lc1).b
Case ":LastDC", ":Last"
LineSC(lc1).R = LastDC(lc1).R
LineSC(lc1).G = LastDC(lc1).G
LineSC(lc1).b = LastDC(lc1).b
Case Else
LineSC(lc1).R = GetRGB(Text_Equation(Ary(((lc1 - 1) * 11) + 3)), "R")
LineSC(lc1).G = GetRGB(Text_Equation(Ary(((lc1 - 1) * 11) + 3)), "G")
LineSC(lc1).b = GetRGB(Text_Equation(Ary(((lc1 - 1) * 11) + 3)), "B")
End Select
End If
If IsNumeric(Ary(((lc1 - 1) * 11) + 4)) = True Then
LineDC(lc1).R = GetRGB(Ary(((lc1 - 1) * 11) + 4), "R")
LineDC(lc1).G = GetRGB(Ary(((lc1 - 1) * 11) + 4), "G")
LineDC(lc1).b = GetRGB(Ary(((lc1 - 1) * 11) + 4), "B")
Else
Select Case Ary(((lc1 - 1) * 11) + 4)
Case ":Random"
LineDC(lc1).R = RndNum(255)
LineDC(lc1).G = RndNum(255)
LineDC(lc1).b = RndNum(255)
Case ":LastSC"
LineDC(lc1).R = LastSC(lc1).R
LineDC(lc1).G = LastSC(lc1).G
LineDC(lc1).b = LastSC(lc1).b
Case ":LastDC", ":Last"
LineDC(lc1).R = LastDC(lc1).R
LineDC(lc1).G = LastDC(lc1).G
LineDC(lc1).b = LastDC(lc1).b
Case Else
LineDC(lc1).R = GetRGB(Text_Equation(Ary(((lc1 - 1) * 11) + 4)), "R")
LineDC(lc1).G = GetRGB(Text_Equation(Ary(((lc1 - 1) * 11) + 4)), "G")
LineDC(lc1).b = GetRGB(Text_Equation(Ary(((lc1 - 1) * 11) + 4)), "B")
End Select
End If
LineType(lc1) = Ary(((lc1 - 1) * 11) + 5)
If IsNumeric(Ary(((lc1 - 1) * 11) + 6)) = True Then
lines(lc1).x1 = Ary(((lc1 - 1) * 11) + 6)
Else
Select Case Ary(((lc1 - 1) * 11) + 6)
Case ":Random"
lines(lc1).x1 = RndNum(DivBy) / DivBy
Case ":LastS"
lines(lc1).x1 = LastS(lc1).x1
Case ":LastD", ":Last"
lines(lc1).x1 = LastD(lc1).x1
Case Else
lines(lc1).x1 = Text_Equation(Ary(((lc1 - 1) * 11) + 6))
End Select
End If
If IsNumeric(Ary(((lc1 - 1) * 11) + 7)) = True Then
lines(lc1).y1 = Ary(((lc1 - 1) * 11) + 7)
Else
Select Case Ary(((lc1 - 1) * 11) + 7)
Case ":Random"
lines(lc1).y1 = RndNum(DivBy) / DivBy
Case ":LastS"
lines(lc1).y1 = LastS(lc1).y1
Case ":LastD", ":Last"
lines(lc1).y1 = LastD(lc1).y1
Case Else
lines(lc1).y1 = Text_Equation(Ary(((lc1 - 1) * 11) + 7))
End Select
End If
If IsNumeric(Ary(((lc1 - 1) * 11) + 8)) = True Then
lines(lc1).x2 = Ary(((lc1 - 1) * 11) + 8)
Else
Select Case Ary(((lc1 - 1) * 11) + 8)
Case ":Random"
lines(lc1).x2 = RndNum(DivBy) / DivBy
Case ":LastS"
lines(lc1).x2 = LastS(lc1).x2
Case ":LastD", ":Last"
lines(lc1).x2 = LastD(lc1).x2
Case Else
lines(lc1).x2 = Text_Equation(Ary(((lc1 - 1) * 11) + 8))
End Select
End If
If IsNumeric(Ary(((lc1 - 1) * 11) + 9)) = True Then
lines(lc1).y2 = Ary(((lc1 - 1) * 11) + 9)
Else
Select Case Ary(((lc1 - 1) * 11) + 9)
Case ":Random"
lines(lc1).y2 = RndNum(DivBy) / DivBy
Case ":LastS"
lines(lc1).y2 = LastS(lc1).y2
Case ":LastD", ":Last"
lines(lc1).y2 = LastD(lc1).y2
Case Else
lines(lc1).y2 = Text_Equation(Ary(((lc1 - 1) * 11) + 9))
End Select
End If
If IsNumeric(Ary(((lc1 - 1) * 11) + 10)) = True Then
LineD(lc1).x1 = Ary(((lc1 - 1) * 11) + 10)
Else
Select Case Ary(((lc1 - 1) * 11) + 10)
Case ":Random"
LineD(lc1).x1 = RndNum(DivBy) / DivBy
Case ":LastS"
LineD(lc1).x1 = LastS(lc1).x1
Case ":LastD", ":Last"
LineD(lc1).x1 = LastD(lc1).x1
Case Else
LineD(lc1).x1 = Text_Equation(Ary(((lc1 - 1) * 11) + 10))
End Select
End If
If IsNumeric(Ary(((lc1 - 1) * 11) + 11)) = True Then
LineD(lc1).y1 = Ary(((lc1 - 1) * 11) + 11)
Else
Select Case Ary(((lc1 - 1) * 11) + 11)
Case ":Random"
LineD(lc1).y1 = RndNum(DivBy) / DivBy
Case ":LastS"
LineD(lc1).y1 = LastS(lc1).y1
Case ":LastD", ":Last"
LineD(lc1).y1 = LastD(lc1).y1
Case Else
LineD(lc1).y1 = Text_Equation(Ary(((lc1 - 1) * 11) + 11))
End Select
End If
If IsNumeric(Ary(((lc1 - 1) * 11) + 12)) = True Then
LineD(lc1).x2 = Ary(((lc1 - 1) * 11) + 12)
Else
Select Case Ary(((lc1 - 1) * 11) + 12)
Case ":Random"
LineD(lc1).x2 = RndNum(DivBy) / DivBy
Case ":LastS"
LineD(lc1).x2 = LastS(lc1).x2
Case ":LastD", ":Last"
LineD(lc1).x2 = LastD(lc1).x2
Case Else
LineD(lc1).x2 = Text_Equation(Ary(((lc1 - 1) * 11) + 12))
End Select
End If
If IsNumeric(Ary(((lc1 - 1) * 11) + 13)) = True Then
LineD(lc1).y2 = Ary(((lc1 - 1) * 11) + 13)
Else
Select Case Ary(((lc1 - 1) * 11) + 13)
Case ":Random"
LineD(lc1).y2 = RndNum(DivBy) / DivBy
Case ":LastS"
LineD(lc1).y2 = LastS(lc1).y2
Case ":LastD", ":Last"
LineD(lc1).y2 = LastD(lc1).y2
Case Else
LineD(lc1).y2 = Text_Equation(Ary(((lc1 - 1) * 11) + 13))
End Select
End If
Next lc1
ReDim LastS(UBound(lines))
ReDim LastD(UBound(LineD))
For lc1 = 1 To UBound(lines)
LastS(lc1).x1 = lines(lc1).x1
LastS(lc1).y1 = lines(lc1).y1
LastS(lc1).x2 = lines(lc1).x2
LastS(lc1).y2 = lines(lc1).y2
LastD(lc1).x1 = LineD(lc1).x1
LastD(lc1).y1 = LineD(lc1).y1
LastD(lc1).x2 = LineD(lc1).x2
LastD(lc1).y2 = LineD(lc1).y2
LastSC(lc1).R = LineSC(lc1).R
LastSC(lc1).G = LineSC(lc1).G
LastSC(lc1).b = LineSC(lc1).b
LastDC(lc1).R = LineDC(lc1).R
LastDC(lc1).G = LineDC(lc1).G
LastDC(lc1).b = LineDC(lc1).b
Next lc1
For lc1 = 1 To ((UBound(Ary) - 2) / 11)
lines(lc1).x1 = lines(lc1).x1 * W1
lines(lc1).y1 = lines(lc1).y1 * H1
lines(lc1).x2 = lines(lc1).x2 * W1
lines(lc1).y2 = lines(lc1).y2 * H1
LineD(lc1).x1 = LineD(lc1).x1 * W1
LineD(lc1).y1 = LineD(lc1).y1 * H1
LineD(lc1).x2 = LineD(lc1).x2 * W1
LineD(lc1).y2 = LineD(lc1).y2 * H1
Next lc1
lc1 = 0
Do
For lc2 = 1 To ((UBound(Ary) - 2) / 11)
Select Case LineType(lc2)
Case ":Line"
thefrm.Line (lines(lc2).x1 + ((LineD(lc2).x1 - lines(lc2).x1) * lc1), lines(lc2).y1 + ((LineD(lc2).y1 - lines(lc2).y1) * lc1))-(lines(lc2).x2 + ((LineD(lc2).x2 - lines(lc2).x2) * lc1), lines(lc2).y2 + ((LineD(lc2).y2 - lines(lc2).y2) * lc1)), RGB(Int(LineSC(lc2).R + ((LineDC(lc2).R - LineSC(lc2).R) * lc1)), Int(LineSC(lc2).G + ((LineDC(lc2).G - LineSC(lc2).G) * lc1)), Int(LineSC(lc2).b + ((LineDC(lc2).b - LineSC(lc2).b) * lc1)))
Case ":Box"
thefrm.Line (lines(lc2).x1 + ((LineD(lc2).x1 - lines(lc2).x1) * lc1), lines(lc2).y1 + ((LineD(lc2).y1 - lines(lc2).y1) * lc1))-(lines(lc2).x2 + ((LineD(lc2).x2 - lines(lc2).x2) * lc1), lines(lc2).y2 + ((LineD(lc2).y2 - lines(lc2).y2) * lc1)), RGB(Int(LineSC(lc2).R + ((LineDC(lc2).R - LineSC(lc2).R) * lc1)), Int(LineSC(lc2).G + ((LineDC(lc2).G - LineSC(lc2).G) * lc1)), Int(LineSC(lc2).b + ((LineDC(lc2).b - LineSC(lc2).b) * lc1))), B
Case ":FilledBox"
thefrm.Line (lines(lc2).x1 + ((LineD(lc2).x1 - lines(lc2).x1) * lc1), lines(lc2).y1 + ((LineD(lc2).y1 - lines(lc2).y1) * lc1))-(lines(lc2).x2 + ((LineD(lc2).x2 - lines(lc2).x2) * lc1), lines(lc2).y2 + ((LineD(lc2).y2 - lines(lc2).y2) * lc1)), RGB(Int(LineSC(lc2).R + ((LineDC(lc2).R - LineSC(lc2).R) * lc1)), Int(LineSC(lc2).G + ((LineDC(lc2).G - LineSC(lc2).G) * lc1)), Int(LineSC(lc2).b + ((LineDC(lc2).b - LineSC(lc2).b) * lc1))), BF
End Select
Next lc2
DoEvents
If lc1 = 1 Then Exit Do
lc1 = lc1 + (1 / LineSpeed)
If lc1 > 1 Then lc1 = 1
Loop
'Ary(1) = LineSpeed
'Ary(2) = DivBy
'Ary(3) = Line1 LineSC
'Ary(4) = Line1 LineDC
'Ary(5) = Line1 LineType
'Ary(6) = Line1 LineSX1
'Ary(7) = Line1 LineSY1
'Ary(8) = Line1 LineSX2
'Ary(9) = Line1 LineSY2
'Ary(10) = Line1 LineDX1
'Ary(11) = Line1 LineDY1
'Ary(12) = Line1 LineDX2
'Ary(13) = Line1 LineDY2

'Ary(14) = Line2 LineSC
End Sub

Sub Form_Implode(thefrm As Object, HowFast As Double)
'-boom
thefrm.Hide
Dim lc1 As Double, FinalX As Long, FinalY As Long
Dim FinalSX As Long, FinalSY As Long
FinalX = thefrm.Width
FinalY = thefrm.Height
FinalSX = thefrm.Left
FinalSY = thefrm.Top
thefrm.Top = thefrm.Top + (thefrm.Height / 2)
thefrm.Left = thefrm.Left + (thefrm.Width / 2)
thefrm.Width = 0
thefrm.Height = 0
thefrm.Show
For lc1 = 0.5 To 1 Step HowFast
thefrm.Width = lc1 * FinalX
thefrm.Height = lc1 * FinalY
thefrm.Left = FinalSX + ((1 - lc1) * (FinalX / 2))
thefrm.Top = FinalSY + ((1 - lc1) * (FinalY / 2))
DoEvents
Next lc1
thefrm.Width = FinalX
thefrm.Height = FinalY
thefrm.Left = FinalSX
thefrm.Top = FinalSY
End Sub

Sub Form_SizeScreen(whichform As Object)
whichform.Left = 0
whichform.Top = 0
whichform.Width = Screen.Width
whichform.Height = Screen.Height
End Sub

Function GetRGB(ByVal Num, ByVal RGorB)
'gets red, green, or blue value of a color integer.  ex: cred = getrgb(form1.backcolor, "r")
Dim Red, Green, Blue, curval
curval = 0
If Num >= 256 ^ 2 Then
curval = Num_RoundDown(val(Num) / (256 ^ 2))
Num = Num - ((256 ^ 2) * curval)
End If
Blue = val(curval)
curval = 0
If val(Num) >= 256 Then
curval = Num_RoundDown(Num / 256)
Num = Num - (256 * curval)
End If
Green = val(curval)
curval = 0
Red = val(Num)
Select Case UCase(Left(RGorB, 1))
Case "R"
GetRGB = Red
Case "G"
GetRGB = Green
Case "B"
GetRGB = Blue
End Select
End Function





Function MakeByteVal(Num As Long)
Do Until Num >= 0 And Num <= 255
If Num < 0 Then Num = Num + 256
If Num > 255 Then Num = Num - 256
Loop
MakeByteVal = Num
End Function

Sub MakeDirectory(DirName As String)
MkDir DirName
End Sub

Function FileOpen(ByVal filename As String)
'i have to comment every damn function!  er, returns contents of file
Dim FileHold, Num
Num = FreeFile
Open filename For Input As Num
Do Until EOF(Num) = True
Line Input #(Num), FileHold
FileOpen = FileOpen & FileHold & Chr$(13) & Chr$(10)
DoEvents
Loop
Close #Num
End Function


Sub FileSave(ByVal filename As String, FileData)
'savesfileblabla
Dim Num
Num = FreeFile
Open filename For Output As Num
Print #Num, FileData
DoEvents
Close #Num
End Sub


Function RemEx(ByVal Num)
'stupid vb, [or me, if theres a way to get by this], for big numbers, puts E in the numbers,
'which confuses itself when it finds that E again.
On Error Resume Next
Dim lc, Exp, DecP
RemEx = Num
If InStr(UCase(Num), "E") = 0 Then Exit Function
Exp = Mid(Num, InStr(UCase(Num), "E") + 1)
Num = Left(Num, InStr(UCase(Num), "E") - 1)
DecP = Len(Num) - InStr(Num, ".")
If Exp < 0 Then
Num = Left(Num, Len(Num) + Exp)
ElseIf Exp > 0 Then
If DecP < Exp Then
Num = Left(Num, Len(Num) - (Exp - DecP))
End If
End If
If IsNumeric(Num) = False Then Num = 0
RemEx = Num * (10 ^ Exp)
End Function

Function Angle_Rotate(X, Y, step) As T_Point
'rotates a 2d angle
Dim length
length = Sqr(X ^ 2 + Y ^ 2)
Angle_Rotate.X = Sin(Angle_GetRadians(X, Y) + step) * length
Angle_Rotate.Y = Cos(Angle_GetRadians(X, Y) + step) * length
End Function

Function Rotate3DLine(startval As T_3DLineCoordinates, VAngle As T_3DPoint) As T_3DLineCoordinates
' i think the same thing as angle rotate, but rotating 2 points of a line
Dim length, newrads
length = Sqr((startval.y1 ^ 2) + (startval.z1 ^ 2))
newrads = Angle_GetRadians(startval.y1, startval.z1) + VAngle.X
Rotate3DLine.y1 = Sin(newrads) * length
Rotate3DLine.z1 = Cos(newrads) * length
length = Sqr((startval.x1 ^ 2) + (Rotate3DLine.z1 ^ 2))
newrads = Angle_GetRadians(startval.x1, Rotate3DLine.z1) + VAngle.Y
Rotate3DLine.x1 = Sin(newrads) * length
Rotate3DLine.z1 = Cos(newrads) * length
length = Sqr((Rotate3DLine.x1 ^ 2) + (Rotate3DLine.y1 ^ 2))
newrads = Angle_GetRadians(Rotate3DLine.x1, Rotate3DLine.y1) + VAngle.z
Rotate3DLine.x1 = Sin(newrads) * length
Rotate3DLine.y1 = Cos(newrads) * length
length = Sqr((startval.y2 ^ 2) + (startval.z2 ^ 2))
newrads = Angle_GetRadians(startval.y2, startval.z2) + VAngle.X
Rotate3DLine.y2 = Sin(newrads) * length
Rotate3DLine.z2 = Cos(newrads) * length
length = Sqr((startval.x2 ^ 2) + (Rotate3DLine.z2 ^ 2))
newrads = Angle_GetRadians(startval.x2, Rotate3DLine.z2) + VAngle.Y
Rotate3DLine.x2 = Sin(newrads) * length
Rotate3DLine.z2 = Cos(newrads) * length
length = Sqr((Rotate3DLine.x2 ^ 2) + (Rotate3DLine.y2 ^ 2))
newrads = Angle_GetRadians(Rotate3DLine.x2, Rotate3DLine.y2) + VAngle.z
Rotate3DLine.x2 = Sin(newrads) * length
Rotate3DLine.y2 = Cos(newrads) * length
End Function

Function Text_Compress1(txt As String, compress As Boolean)
'obviously never finished
Dim Letters As String, lc As Long, base As Integer, txtbin As String
For lc = 1 To Len(txt)
If InStr(Letters, Mid$(txt, lc, 1)) = 0 Then Letters = Letters & Mid$(txt, lc, 1)
Next lc

For lc = 8 To 1
base = lc
If 2 ^ lc < Len(Letters) Then Exit For
Next lc










End Function

Function Text_MGFencrypt5(ByVal txt As String, ByVal Pass As String, Encrypt As Boolean) As String
'does it work?  you tell muh!
Dim ASCdif() As Long
Dim lc1 As Long
Dim lc2 As Long
Dim Dif As Long
ReDim ASCdif(Len(Pass))
For lc1 = 1 To UBound(ASCdif)
Dif = 0
For lc2 = 1 To lc1 - 1
Dif = Dif - Asc(Mid$(Pass, lc2, 1))
Next lc2
For lc2 = lc1 + 1 To UBound(ASCdif)
Dif = Dif + Asc(Mid$(Pass, lc2, 1))
Next lc2
ASCdif(lc1) = Dif
Next lc1
Dif = 0
For lc1 = 1 To Len(txt)
Dif = Dif + 1
If Dif >= Len(Pass) Then Dif = 1
If Encrypt = True Then Text_MGFencrypt5 = Text_MGFencrypt5 & Chr(MakeByteVal(Asc(Mid$(txt, lc1, 1)) + ASCdif(Dif)))
If Encrypt = False Then Text_MGFencrypt5 = Text_MGFencrypt5 & Chr$(MakeByteVal(Asc(Mid$(txt, lc1, 1)) - ASCdif(Dif)))
Next lc1
End Function



Sub Pause(tm As Double)
'pauses program code exectution for tm seconds.
Dim st As String: DoEvents: st = Timer
Do Until Timer >= st + tm
DoEvents
Loop
End Sub
Function PixGet(thefrm As Form, X, Y)
'get pixel color
Dim wind
DoEvents
wind = GetDC(thefrm.hwnd)
PixGet = GetPixel(wind, X, Y)
End Function

Sub PixSet(thefrm As Form, X, Y, clr)
'set pixel color
Dim wind, returnval
DoEvents
wind = GetWindowDC(thefrm.hwnd)
returnval = SetPixel(wind, X, Y, clr)
End Sub

Function RndNum(Max)
'raanndom numbah
'Pause (0.01)
Randomize ' Timer
RndNum = Int((Max + 1) * Rnd + 0)

End Function

Sub SCRclose()
End
End Sub

Sub SCRload()
'for loading screensaver apps - kinda buggy with winxp, dunno why.
If Command$ = "" Then GoTo SetShow
Select Case LCase(Left$(Command$, 2))
Case "/p"
'preview
End
Case "/c"
SetShow:
Form1.Show
'settings / configure
Case "/s"
SCRopen
Form1.Show
'run screensaver
End Select
End Sub

Sub SCRopen()
If App.PrevInstance = True Then End
End Sub

Sub Text_FillArrayRndNums(Ary() As Long, Optional Max As Variant)
If IsMissing(Max) = True Then
Max = UBound(Ary)
End If
ReDim Ary(Max)
Dim CurNum As Long, EmptySlotCount As Long, EmptySlotNum As Long
Dim CurEmptySlot As Long, RndN As Long, CurSlot As Long, CurSlotCount As Long
Dim CurSlotLC1 As Long
For CurNum = 1 To Max
EmptySlotNum = 0
RndN = 0
CurSlotCount = 0
EmptySlotNum = Max - CurNum + 1
Randomize Timer
RndN = Int(Rnd * EmptySlotNum + 1)
For CurSlotLC1 = 1 To Max
If val(Ary(CurSlotLC1)) = 0 Then CurSlotCount = CurSlotCount + 1
If CurSlotCount = RndN Then
CurSlot = CurSlotLC1
Exit For
End If
Next CurSlotLC1
Ary(CurSlot) = CurNum
Next CurNum
End Sub


Function Text_MGFencrypt4(ByVal txt As String, ByVal Pass As String, ByVal base As Integer) As String
'good, reliable encryption.  Hard to break cuza password, but if you had the decrypted text, pass, base, and encrypted text, you could figure it out.
'encrypting function.  basically, to encrypt, function adds asc of current character to (asc of cooresponding character in pass * base)
'to use, encrypt using any password (string) and any base (integer).  to decrypt, use same pass, but the negative value of the encrypting value.
Dim lc1 As Integer, lc2 As Integer, lc3 As Integer, hold() As Integer
ReDim hold(Len(txt))
lc2 = 0
For lc1 = 1 To Len(txt)
If lc2 >= Len(Pass) Then lc2 = 0
lc2 = lc2 + 1
hold(lc1) = Asc(Mid$(txt, lc1, 1)) + (Asc(Mid$(Pass, lc2, 1)) * base)
Do Until hold(lc1) >= 1 And hold(lc1) <= 255
If hold(lc1) > 255 Then hold(lc1) = hold(lc1) - 255
If hold(lc1) < 1 Then hold(lc1) = hold(lc1) + 255
Loop
Next lc1
For lc1 = 1 To Len(txt)
Text_MGFencrypt4 = Text_MGFencrypt4 & Chr$(hold(lc1))
Next lc1
End Function




Function Text_MGFencrypt3(ByVal txt As String, ByVal base As Integer, ByVal StartASC As String, ByVal Encrypt As Boolean) As String
'decent encryption.  only setback - the returned encrypted code will most likely be over twice as long.
If base <= 1 Or base >= 10 Then
Text_MGFencrypt3 = "Error - Base must be between 1 And 10."
Exit Function
End If
'97 - 122
Dim Alph() As String
Dim lc1 As Long, lc2 As Long, lc3 As Long, Exp As Integer
Dim Bcode() As String
Dim hold As String
If Encrypt = True Then
lc1 = 0
Do Until base ^ lc1 >= 255
lc1 = lc1 + 1
Loop
Exp = lc1
lc2 = -1
ReDim Alph(base - 1)
For lc1 = 1 To 26
lc2 = lc2 + 1
If lc2 >= base Then
lc2 = 0
End If
Alph(lc2) = Alph(lc2) & Chr$((Asc(StartASC) - 1) + lc1)
Next lc1
For lc1 = 0 To base - 1
Next lc1
lc1 = 0
ReDim Bcode(base ^ Exp)
For lc1 = 1 To Exp
Bcode(1) = Bcode(1) & "0"
Next lc1
For lc1 = 2 To base ^ Exp
Bcode(lc1) = Left$(Bcode(lc1 - 1), Len(Bcode(lc1 - 1)) - 1) & val(Right$(Bcode(lc1 - 1), 1)) + 1
For lc2 = Exp To 1 Step -1
If val(Mid$(Bcode(lc1), lc2, 1)) >= base Then
Bcode(lc1) = Left$(Bcode(lc1), lc2 - 2) & val(Mid$(Bcode(lc1), lc2 - 1, 1)) + 1 & "0" & Mid$(Bcode(lc1), lc2 + 1)
End If
Next lc2
Next lc1
For lc1 = 1 To Len(txt)
Text_MGFencrypt3 = Text_MGFencrypt3 & Bcode(Asc(Mid$(txt, lc1, 1)))
Next lc1
hold = ""
For lc1 = 1 To Len(Text_MGFencrypt3)
Randomize Timer
hold = hold & Mid$(Alph(val(Mid$(Text_MGFencrypt3, lc1, 1))), Int(Rnd * Len(Alph(val(Mid$(Text_MGFencrypt3, lc1, 1)))) + 1), 1)
Next lc1
Text_MGFencrypt3 = hold

ElseIf Encrypt = False Then
Exp = 0
Do Until base ^ Exp >= 255
Exp = Exp + 1
Loop
lc2 = -1
ReDim Alph(base - 1)
For lc1 = 1 To 26
lc2 = lc2 + 1
If lc2 >= base Then
lc2 = 0
End If
Alph(lc2) = Alph(lc2) & Chr$((Asc(StartASC) - 1) + lc1)
Next lc1
For lc1 = 0 To base - 1
Next lc1
For lc1 = 1 To Len(txt)
For lc2 = 0 To base - 1
If Text_NumTextExists(Alph(lc2), Mid$(txt, lc1, 1)) > 0 Then
hold = hold & lc2
Exit For
End If
Next lc2
Next lc1
lc1 = 0
ReDim Bcode(base ^ Exp)
For lc1 = 1 To Exp
Bcode(1) = Bcode(1) & "0"
Next lc1
For lc1 = 2 To base ^ Exp
Bcode(lc1) = Left$(Bcode(lc1 - 1), Len(Bcode(lc1 - 1)) - 1) & val(Right$(Bcode(lc1 - 1), 1)) + 1
For lc2 = Exp To 1 Step -1
If val(Mid$(Bcode(lc1), lc2, 1)) >= base Then
Bcode(lc1) = Left$(Bcode(lc1), lc2 - 2) & val(Mid$(Bcode(lc1), lc2 - 1, 1)) + 1 & "0" & Mid$(Bcode(lc1), lc2 + 1)
End If
Next lc2
Next lc1
For lc1 = 1 To Len(txt) / Exp
For lc2 = 1 To base ^ Exp
If Mid$(hold, (lc1 * Exp) - Exp + 1, Exp) = Bcode(lc2) Then
Text_MGFencrypt3 = Text_MGFencrypt3 & Chr$(lc2)
Exit For
End If
Next lc2
Next lc1
End If
End Function


Function Text_MGFscramble1(ByVal txt As String) As String
Dim Final() As Long
Dim lc1 As Long
ReDim Final(Len(txt))
Text_FillArrayRndNums Final()
For lc1 = 1 To Len(txt)
Text_MGFscramble1 = Text_MGFscramble1 & Mid$(txt, Final(lc1), 1)
Next lc1
End Function
Function Text_MGFscramble2(ByVal txt As String, Optional WordSep As Variant) As String
If IsMissing(WordSep) = True Then WordSep = " "
Dim lc1 As Long
Dim wrds() As String
ReDim wrds(1)
Dim lc2 As Long
lc2 = 1
For lc1 = 1 To Len(txt)
If Mid$(txt, lc1, Len(WordSep)) = WordSep Then
lc2 = lc2 + 1
ReDim Preserve wrds(lc2)
lc1 = lc1 + Len(WordSep)
End If
wrds(lc2) = wrds(lc2) & Mid$(txt, lc1, 1)
Next lc1
For lc1 = 1 To lc2
wrds(lc1) = Text_MGFscramble1(wrds(lc1))
Next lc1
For lc1 = 1 To lc2
Text_MGFscramble2 = Text_MGFscramble2 & WordSep & wrds(lc1)
Next lc1
Text_MGFscramble2 = Mid$(Text_MGFscramble2, Len(WordSep) + 1)
End Function




Function Text_ReplaceText(Text As String, Rep1 As String, Rep2 As String) As String
'word, i used to use this function to make a program for editing vb code, like if i needed to change a variable name.
Dim CurNum2 As Integer
For CurNum2 = 1 To (Len(Text))
If Mid$(Text, CurNum2, Len(Rep1)) = Rep1 Then
Text_ReplaceText = Text_ReplaceText & Rep2
CurNum2 = CurNum2 + Len(Rep1) - 1
Else
Text_ReplaceText = Text_ReplaceText & Mid$(Text, CurNum2, 1)
End If
Next CurNum2
End Function




Function Text_NumTextExists(txt, Search) As Integer
'oh my, was this before i learned about instr?
Dim CurNum As Integer
Text_NumTextExists = 0
For CurNum = 1 To (Len(txt) - Len(Search) + 1)
If Mid$(txt, CurNum, Len(Search)) = Search Then
Text_NumTextExists = Text_NumTextExists + 1
CurNum = CurNum + Len(Search) - 1
End If
Next CurNum
End Function



Function Text_Equation(ByVal Eqt As String)
'equation parser (ex: "5+(6^2)/((2+1)*6)" returns "7")
Dim NNL As String, lc1 As Long, lc2 As Long, lc3 As Long, Ary()
Dim UB1 As Boolean, PC1 As Long, PC2 As Long, PC3 As Long
Dim OprLst() As String, OprMax As Long
Dim PC4 As Double, UB2 As Boolean
Dim OprPos As Long, NewVal As Double
OprMax = 3
ReDim OprLst(OprMax)
OprLst(1) = "^"
OprLst(2) = "*/"
OprLst(3) = "+-"
NNL = "()^*/+-"
lc2 = 1
ReDim Ary(lc2)
For lc1 = 1 To Len(Eqt)
If Text_NumTextExists(NNL, Mid$(Eqt, lc1, 1)) > 0 And lc1 > 1 Then
lc2 = lc2 + 1
ReDim Preserve Ary(lc2)
Else
UB1 = False
If lc1 > 2 Then
If Mid$(Eqt, lc1 - 1, 1) = "-" And Text_NumTextExists(NNL, Mid$(Eqt, lc1 - 2, 1)) > 0 And Mid$(Eqt, lc1 - 2, 1) <> ")" Then UB1 = True
End If
If lc1 > 1 And UB1 = False Then
If Text_NumTextExists(NNL, Mid$(Eqt, lc1 - 1, 1)) > 0 Then
If (lc1 = 2) And Left$(Eqt, 1) = "-" Then
Else
lc2 = lc2 + 1
ReDim Preserve Ary(lc2)
End If
End If
End If
End If
Ary(lc2) = Ary(lc2) & Mid$(Eqt, lc1, 1)
Next lc1
Do Until UBound(Ary) = 1
StartLoop:
Array_ClearNulls Ary()
PC4 = 0
PC3 = 1
PC2 = 0
PC1 = 0
For lc1 = 1 To UBound(Ary)
Select Case Ary(lc1)
Case "("
PC1 = PC1 + 1
If PC1 > PC2 Then
PC2 = PC1
PC3 = lc1
End If
Case ")"
PC1 = PC1 - 1
End Select
Next lc1
If UBound(Ary) = 1 Then Exit Do
If Ary(PC3) = "(" And Ary(PC3 + 1) <> ")" Then
If Ary(PC3) = "(" And Ary(PC3 + 2) = ")" Then
Ary(PC3) = ""
Ary(PC3 + 2) = ""
GoTo StartLoop
End If
End If
PC4 = UBound(Ary)
For lc1 = PC3 To UBound(Ary)
If Ary(lc1) = ")" Then
PC4 = lc1
Exit For
End If
Next lc1
OprPos = 1
UB2 = False
For lc1 = 1 To OprMax
For lc2 = PC3 To PC4
For lc3 = 1 To Len(OprLst(lc1))
If Ary(lc2) = Mid$(OprLst(lc1), lc3, 1) And UB2 = False Then
UB2 = True
OprPos = lc2
End If
If UB2 = True Then Exit For
Next lc3
If UB2 = True Then Exit For
Next lc2
If UB2 = True Then Exit For
Next lc1
Select Case Ary(OprPos)
Case "^"
NewVal = val(Ary(OprPos - 1)) ^ val(Ary(OprPos + 1))
Case "*"
NewVal = val(Ary(OprPos - 1)) * val(Ary(OprPos + 1))
Case "/"
NewVal = val(Ary(OprPos - 1)) / val(Ary(OprPos + 1))
Case "+"
NewVal = val(Ary(OprPos - 1)) + val(Ary(OprPos + 1))
Case "-"
NewVal = val(Ary(OprPos - 1)) - val(Ary(OprPos + 1))
Case Else
NewVal = 0
End Select
Ary(OprPos - 1) = ""
Ary(OprPos) = NewVal
Ary(OprPos + 1) = ""
Array_ClearNulls Ary()
Loop
Text_Equation = Ary(1)
End Function

Function Num_RoundDown(Num)
Num_RoundDown = Num
Dim Trim1 As Long
For Trim1 = 1 To Len(Num)
If Mid(Num, Trim1, 1) = "." Then
Num_RoundDown = Left(Num, Trim1 - 1)
Exit For
End If
Next Trim1
End Function
Function Num_RoundUp(Num)
If InStr(Str(Num), ".") > 0 Then
Num_RoundUp = val(Mid$(Str(Num), 1, InStr(Str(Num), ".") - 1)) + 1
Else
Num_RoundUp = Num
End If
End Function
Function Num_AddZeros(Num, Zeros)
'to keep fast changing numbers neat ;]
Dim X
For X = 1 To Zeros - Len(Num)
Num_AddZeros = "0" & Num_AddZeros
Next X
Num_AddZeros = Num_AddZeros & Num
End Function


Sub UpdateDisplay()
'the main 3d drawing for my 3dline program
Dim lc, NewPos As T_3DLineCoordinates
For lc = 1 To UBound(LinePos)





NewPos.x1 = LinePos(lc).x1
NewPos.y1 = LinePos(lc).y1
NewPos.z1 = LinePos(lc).z1
NewPos.x2 = LinePos(lc).x2
NewPos.y2 = LinePos(lc).y2
NewPos.z2 = LinePos(lc).z2

NewPos = Rotate3DLine(NewPos, RotateAllAngle)

NewPos.x1 = NewPos.x1 - ViewPoint.X
NewPos.y1 = NewPos.y1 - ViewPoint.Y
NewPos.z1 = NewPos.z1 - ViewPoint.z
NewPos.x2 = NewPos.x2 - ViewPoint.X
NewPos.y2 = NewPos.y2 - ViewPoint.Y
NewPos.z2 = NewPos.z2 - ViewPoint.z
'AD NewPos.X1
'  NewPos = Rotate3DLine(NewPos, ViewAngle)
'AD NewPos.X1
'AD NewPos.Y1
'AD ""
'AD Angle_GetRadians(NewPos.X1, NewPos.Y1) / ViewWindow.X
'Form1.LineShow(lc).X1 = (Form1.Width / 2) + ((Angle_GetRadians(NewPos.X1, NewPos.Y1) / ViewWindow.x) * (Form1.Width / 2))
'Form1.LineShow(lc).Y1 = (Form1.Height / 2) - ((Angle_GetRadians(NewPos.Z1, NewPos.Y1) / ViewWindow.Z) * (Form1.Height / 2))
'AD Form1.LineShow(lc).X1 / Form1.Width
'AD Form1.LineShow(lc).Y1 / Form1.Height

'Form1.LineShow(lc).X2 = (Form1.Width / 2) + ((Angle_GetRadians(NewPos.X2, NewPos.Y2) / ViewWindow.x) * (Form1.Width / 2))
'Form1.LineShow(lc).Y2 = (Form1.Height / 2) - ((Angle_GetRadians(NewPos.Z2, NewPos.Y2) / ViewWindow.Z) * (Form1.Height / 2))

Next lc



'If ZoomFact = 0 Then ZoomFact = 1
'With Form1
'Dim lc, rval1 As T_3DPoint, rval2 As T_3DPoint
'For lc = 1 To UBound(LinePos)
'rval1 = Angle_Rotate3D(LinePos(lc).X1, LinePos(lc).Y1, LinePos(lc).Z1, Xang, Yang, Zang)
'rval2 = Angle_Rotate3D(LinePos(lc).X2, LinePos(lc).Y2, LinePos(lc).Z2, Xang, Yang, Zang)
'rval1.X = rval1.X * ZoomFact
'rval1.Y = rval1.Y * ZoomFact
'rval1.Z = rval1.Z * ZoomFact
'rval2.X = rval2.X * ZoomFact
'rval2.Y = rval2.Y * ZoomFact
'rval2.Z = rval2.Z * ZoomFact

'DistSizeFilter rval1.X, rval1.Y, rval1.Z, DistSizeMax
'DistSizeFilter rval2.X, rval2.Y, rval2.Z, DistSizeMax

'.LineShow(lc).X1 = (Form1.ScaleWidth / 2) + rval1.X
'.LineShow(lc).Y1 = (Form1.ScaleHeight / 2) - rval1.Z
'.LineShow(lc).X2 = (Form1.ScaleWidth / 2) + rval2.X
'.LineShow(lc).Y2 = (Form1.ScaleHeight / 2) - rval2.Z
'Next lc
'End With
End Sub

Function Angle_Rotate3D(Xval, Yval, Zval, xangle, yangle, zangle) As T_3DPoint
'rotates a 3d angle
Dim length, newrads
length = Sqr((Yval ^ 2) + (Zval ^ 2))
newrads = Angle_GetRadians(Yval, Zval) + xangle
Angle_Rotate3D.Y = Sin(newrads) * length
Angle_Rotate3D.z = Cos(newrads) * length
length = Sqr((Xval ^ 2) + (Angle_Rotate3D.z ^ 2))
newrads = Angle_GetRadians(Xval, Angle_Rotate3D.z) + yangle
Angle_Rotate3D.X = Sin(newrads) * length
Angle_Rotate3D.z = Cos(newrads) * length
length = Sqr((Angle_Rotate3D.X ^ 2) + (Angle_Rotate3D.Y ^ 2))
newrads = Angle_GetRadians(Angle_Rotate3D.X, Angle_Rotate3D.Y) + zangle
Angle_Rotate3D.X = Sin(newrads) * length
Angle_Rotate3D.Y = Cos(newrads) * length
End Function
