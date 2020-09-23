VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000006&
   Caption         =   "LineSpec3D bare"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   593
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'LineSpec3D bare 1.02  written by easfw - fluoats@hotmail.com

'This code is a stripped-down version of LineSpec3D

'I am releasing this code because it is simple to use
'and still fast.

'At run-time, press the number keys 1 thru 4 ..


'Here is how easy it is to build a model:
 '1. Define at least 2 points
 '   a. Call PointDEF(Rnd - 0.5, Rnd - 0.5, Rnd - 0.5)
 '   b. Call PointDEF(Rnd - 0.5, Rnd - 0.5, Rnd - 0.5)
 '2. Define a line
 '   a. Call LineDEF(1,2)   meaning (point 1, point 2)
 
'Where do you type this?
'Private Sub LineSpec()


'Brief description of other subs:
 'AnimWireframe
  'writes a background to 2D array backbuff()
  'animates the model
  
 'antialias
  'generates an alpha for pixels along a line's path
  'as a Single between 0 and 1
  
 'ScaleModel
  'I put this call at the end of LineSpec() so that a model
  'with point x,y,z coordinates close to 1, or much higher
  'will be recognizeable within run-time window

 'Form_MouseMove
  'Main purpose is to allow rotation change via mouse -
  'code is rather complicated but when you click and drag,
  'this sub changes the values of axi and ayi, which
  'are read by AnimWireframe.
  'Also needed are Form_MouseDown and Form_MouseUp
  
 'Form_Load is all the way at the bottom
 
  
'From here on, the variables get rather messy;  I've commented
'somewhat heavily.

'I am also working on a version that
'will incorporate a mouse-driven modeller.

  
'These arrays hold 'backbuffer' data used to erase/draw
'300000 pixels quickly
Dim savX(0 To 300000) As Long    'Filled
Dim savy(0 To 300000) As Long    'by
Dim savAlpha(0 To 300000) As Single 'antialias()

Dim csn(0 To 300000) As Boolean 'AnimWireframe / antialias
Dim dsn(0 To 300000) As Boolean
Dim esn(0 To 300000) As Boolean
Dim pixarray As Long

'Line Info arrays -
'increase these if you want more points in any wireframe
Dim LnPt1(1 To 1001) As Integer, LnPt2(1 To 1001) As Integer
Dim px(1 To 1001) As Single, pY(1 To 1001) As Single
'pX and pY are used by AnimWireframe() and antialias()
Dim pointcount As Integer

'increase these if you want more lines
Dim opac(1 To 1000) As Byte 'Line Intensity
Dim cR(1 To 1000) As Byte
Dim cG(1 To 1000) As Byte
Dim cB(1 To 1000) As Byte
Dim ds(1 To 1000) As Byte 'drawstyle
'these arrays are used by LineDEF and AnimWireframe

Dim linecount As Integer

Dim modelcount As Byte
Dim modelselect As Byte 'Subs Form_Keydown and LineSpec()

Dim fw As Long, fh As Long 'formwidth, formheight
Dim fw1 As Long, fw2 As Long
Dim sw As Long, sh As Long
Dim eye As Long, radius As Single
Dim axi As Single, ayi As Single 'model rotation
Dim ay As Single, ax As Single, az As Single

Dim breakloop As Boolean, resized As Boolean
Dim newbackground As Boolean

'Hold model data .. the ! means As Single
Dim x3!(1 To 1000), y3!(0 To 1000), z3!(0 To 1000)
'Hold scaled model data
Dim x5!(1 To 1000), y5!(1 To 1000), z5!(1 To 1000)

Dim sr(0 To 128) As Single 'Form_Load() fills these 'look-up table'
Dim gs(0 To 255) As Single 'arrays which antialias() accesses

Dim drawselect As Byte 'accessed by AnimWireframe()

'"Virtual" controls .. the % means As Integer
Dim vleft%(1 To 1), vtop%(1 To 1)
Dim vright%(1 To 1), vbot%(1 To 1)
Dim vmax(1 To 1) As Single
Dim vmin(1 To 1) As Single
Dim vval(1 To 1) As Single

'Mouse Control
Dim yInit As Integer, selectv As Byte
Dim xr As Integer, yr As Integer
Dim xr2 As Integer, yr2 As Integer
Dim pressed As Boolean
Dim Elap As Long 'timing

'Multi-Purpose
Dim BGR As Long, cL As Long
Dim lngFN&, sngAP!, sngX!, sngY! 'the & means As Long
Dim lngX&, lngY&, intN1%, intN2% 'the % means As Integer
Dim bytW As Byte, sng1 As Single 'the ! means As Single
Dim shape As Byte, pow As Single
Dim ns1 As Long, ns2 As Long, n2 As Long
Const pi As Single = 3.14159265
Const twopi = 2 * pi

Dim Fin As Boolean 'If true then Exits Do in AnimWireframe,
'then Unloads Me

'Background appearance - See Form_Load
Dim bwi As Boolean
Dim wildbackground As Boolean

'Horizontal Background Fade
Dim ditr As Byte, ditg As Byte, ditb As Byte
Dim h As Integer, q As Integer
Dim rr As Byte, gg As Byte, bb As Byte
Dim dr As Byte, dg As Byte, db As Byte


Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub LineSpec()
pointcount = 0: linecount = 0

Select Case modelselect 'modelselect is changed at Form_KeyDown
Case 0 'First model:  The Cube
 Call PointDEF(-0.5, 0.5, -0.5)
 Call PointDEF(0.5, 0.5, -0.5)
 Call PointDEF(-0.5, -0.5, -0.5)
 Call PointDEF(0.5, -0.5, -0.5)
 Call PointDEF(-0.5, 0.5, 0.5)
 Call PointDEF(0.5, 0.5, 0.5)
 Call PointDEF(-0.5, -0.5, 0.5)
 Call PointDEF(0.5, -0.5, 0.5)
 
                     'red,green
 Call LineDEF(1, 2, , 255, 230)
 Call LineDEF(1, 3, 1) 'the ", 1" means "look of previous line"
 Call LineDEF(1, 5, 1)
 Call LineDEF(4, 2, , 96, 0, 255, 128) '128 for intensity
 Call LineDEF(4, 3, 1)
 Call LineDEF(4, 8, 1)
 Call LineDEF(7, 3, , 99, 99, 245, 195)
 Call LineDEF(7, 5, 1)
 Call LineDEF(7, 8, 1)
 Call LineDEF(6, 2, , 158, 158, 158, 250)
 Call LineDEF(6, 5, 1)
 Call LineDEF(6, 8, 1)
 
 'Extra line for style
 Call PointDEF(0.3, 0.3, 0.3)
 Call PointDEF(-0.3, 0.3, 0.3)
 Call LineDEF(9, 10, , , 129, , 255) 'A green line at 150 intensity
 
 'Here is one last thing you can do with a line:
 'Call LineDEF(1,2, , red, green, blue, intensity, 0 1 or 2)
 
 'Drawstyle 0, 1, or 2.  Try them out
 
 
'On to a different model:
Case 1 '"The Face" - Type '2' to see this model at run-time

 'upper lip
 Call PointDEF(-0.91, -0.435, -0.5)
 Call PointDEF(-0.84, -0.404, -0.52)
 Call PointDEF(-0.65, -0.34, -0.6)
 Call PointDEF(-0.43, -0.276, -0.64)
 Call PointDEF(-0.19, -0.25, -0.65)
 Call PointDEF(0, -0.28, -0.65)
 Call PointDEF(0.19, -0.25, -0.65)
 Call PointDEF(0.43, -0.276, -0.64)
 Call PointDEF(0.65, -0.34, -0.6)
 Call PointDEF(0.84, -0.404, -0.52)
 Call PointDEF(0.91, -0.435, -0.5)
 Call LineDEF(1, 2, , 255, 120, 255, , 1) 'Drawstyle!  0 1 or 2
 Call LineDEF(2, 3, 1)
 Call LineDEF(3, 4, 1)
 Call LineDEF(4, 5, 1)
 Call LineDEF(5, 6, 1)
 Call LineDEF(6, 7, 1)
 Call LineDEF(7, 8, 1)
 Call LineDEF(8, 9, 1)
 Call LineDEF(9, 10, 1)
 Call LineDEF(10, 11, 1)
 
 'these are the coordinates for the lower lip
 Call PointDEF(-0.88, -0.46, -0.5) '
 Call PointDEF(-0.655, -0.56, -0.56)
 Call PointDEF(-0.35, -0.695, -0.59)
 Call PointDEF(0, -0.723, -0.59)
 Call PointDEF(0.35, -0.695, -0.59)
 Call PointDEF(0.655, -0.56, -0.56)
 Call PointDEF(0.88, -0.46, -0.5)
 Call LineDEF(1, 12, , 255, 120, 255)
 Call LineDEF(12, 13, 1)
 Call LineDEF(13, 14, 1)
 Call LineDEF(14, 15, 1)
 Call LineDEF(15, 16, 1)
 Call LineDEF(16, 17, 1)
 Call LineDEF(17, 18, 1)
 Call LineDEF(18, 11, 1)
 
 'Nose
 Call PointDEF(0, 0.78, -0.57)
 Call PointDEF(0, 0.01, -0.95)
 Call PointDEF(-0.21, -0.09, -0.7)
 Call PointDEF(0.21, -0.09, -0.7)
 Call LineDEF(19, 20, , 255, 176, 0, 192)
 Call LineDEF(20, 21, , 255, 176, 0, , 1)
 Call LineDEF(20, 22, 1)
 
 'This generates the eye circle
 lngX = pointcount: lngY = lngX + 1
 For sngAP = 0 To twopi Step twopi / 23
 lngX = lngX + 1
 Call PointDEF(0.14 * Sin(sngAP) - 0.56, 0.14 * Cos(sngAP) + 1.1, -0.5)
 If lngX < lngY + 23 Then Call LineDEF(lngX, lngX + 1, , 165, 128, 255)
 Next sngAP
 

Case 2 '"Wavy Scape - Type '3' on the keyboard to see this model
 For sngX = -4.5 To 4.5 Step 1
 For sngY = -4.5 To 4.5 Step 1
 Call PointDEF(sngX, sngY, 0.3 * Sin(sngY + sngX))
 Next sngY
 Next sngX
 bytW = 80
 For lngY = 0 To 80 Step 10
 For lngX = 1 To 9 Step 1
  lngFN = lngX + lngY
  Call LineDEF(lngFN, lngFN + 1, , bytW, bytW, 255 - bytW)
  Call LineDEF(lngFN, lngFN + 10, , , 255 - bytW)
  Call LineDEF(lngFN, lngFN + 11, , 255 - bytW, 255 - bytW, 192)
  bytW = bytW + 1
 Next lngX
 Next lngY

Case 3 'Random Lines
 For ns1 = 1 To 200
 Call PointDEF((Rnd - 0.5), (Rnd - 0.5), (Rnd - 0.5))
 Next ns1
 For ns1 = 2 To 200 Step 2
 bytW = ns1
 'Call LineDEF(ns1, ns1 - 1, , bytW + 35, 255, bytW + 35, _
              105 * Rnd + 150, Rnd * 2)
 Call LineDEF(ns1, ns1 - 1, , Rnd * 255, Rnd * 255, Rnd * 255, Rnd * 150 + 105)
 Next ns1

Case 4
 '...
End Select



ScaleModel
'Call PointDEF(x,y,z) writes to x3 y3 z3 arrays.
'ScaleModel multiplies x3 y3 z3 by a certain value and writes to
'x5 y5 z5 arrays, where most models are scaled to fit view surface
'AnimWireframe draws according to what's in the x5 .. arrays.
'If I were to save points to file, I'd record x3 .. values
End Sub

Private Sub Form_Activate()
Form1.Cls 'I need to Cls or it blanks my background later
LineSpec 'produce our model!
Call AnimWireframe
End Sub
Private Sub AnimWireframe()
Static r2 As Integer, g2 As Long, b2 As Long
Static r As Single, g As Single, b As Single
Static a As Single, s1 As Single, s2 As Single
Static intensitybyt As Byte
Static savcL(0 To 100000) As Long
Static savBG(0 To 100000) As Long
Static draw(0 To 100000) As Boolean, bcleansurf(0 To 1600, 0 To 1200) As Boolean
Static bclneapix(0 To 100000) As Boolean
Static svBG(0 To 100000) As Long
Static svX(0 To 100000) As Long
Static svY(0 To 100000) As Long
Static backbuff() As Long, backpxln() As Long

Static x1 As Single, y1 As Single
Static x2 As Single, y2 As Single
Static x4 As Long, y4 As Long
Static X As Single, Y As Single, z As Single
Static vpd As Single  'vanishing-point distortion, or near-far distortion
Static csay As Single, snay As Single 'csay = cos(ay)
Static csaz As Single, snaz As Single
Static csax As Single, snax As Single
Static cswap As Integer, mode2 As Boolean
Static ax1 As Single, ay1 As Single
Static arraylen As Long, realpix As Long, pixeln As Long

'This part is used to initialize a bunch of things
'Some variables can be changed safely - they have comments
If newbackground Then
 Select Case fw 'fw is changed at a Form Resize
 Case Is <= 8: ReDim backbuff(0 To 8, 0 To 6): fw = 8: fh = 6
 Case Is <= 640: ReDim backbuff(0 To 640, 0 To 480): fw = 640: fh = 480: ReDim backpxln(0 To 640, 0 To 480)
 Case Is <= 800: ReDim backbuff(0 To 800, 0 To 600): fw = 800: fh = 600: ReDim backpxln(0 To 800, 0 To 600)
 Case Is <= 1024: ReDim backbuff(0 To 1024, 0 To 768): fw = 1024: fh = 768: ReDim backpxln(0 To 1024, 0 To 768)
 Case Is <= 1280: ReDim backbuff(0 To 1280, 0 To 1024): fw = 1280: fh = 1024: ReDim backpxln(0 To 1280, 0 To 1024)
 Case Is <= 1600: ReDim backbuff(0 To 1600, 0 To 1200): fw = 1600: fh = 1200: ReDim backpxln(0 To 1600, 0 To 1200)
 End Select
 'wildbackground = True
 If Not wildbackground Then
 'background 'dither' color
 ditr = 170: ditg = 190: ditb = 210
 Select Case bwi
 Case True: Form1.BackColor = vbBlack
 Case Else: Form1.BackColor = vbWhite
 End Select
 For h = 255 To 0 Step -1
 q = 741 - h * 5
 dr = ditr / 255 * h
 dg = ditg / 255 * h
 db = ditb / 255 * h
 Select Case bwi
 Case 0: rr = 255 - dr: gg = 255 - dg: bb = 255 - db: q = 741 - h * 3
 Case Else: rr = dr: gg = dg: bb = db: q = 601 - h * 3
 End Select
 BGR = RGB(rr, gg, bb)
 For n2 = 0 To fw Step 1
  For cL = q - 2 To q Step 1
  If cL > -1 And cL < fh + 1 Then backbuff(n2, cL) = BGR
  Next cL
 Next n2
 Next h
 Else 'Else wildbackground = True:  draw this wild pattern
  X = 35 'Rnd * 5 + 8
  Y = 9500 '+ Rnd * 6000
  For sng1 = 0 To fw
  For n2 = 0 To fh
  cL = (12254 * Sin(n2 * sng1 / Y) + 12024 * Sin(sng1 / (X))) / 98
  If cL > 16777216 Then
  cL = 16777216
  ElseIf cL < 0 Then
  cL = 0
  End If
  r2 = cL& And &HFF
  If bwi Then If r2 < 120 Then r2 = 120
  g2 = 255
  b2 = 255 - r2
  'the If here produces a fractal-looking something-or-other
  'If n And n2 Then
  r2 = 255 - r2
  'End If
  If bwi Then
  cL = RGB(255 - r2, 255 - g2, 255 - b2)
  Else
  cL = RGB(r2, g2, b2)
  End If
  backbuff(sng1, n2) = cL&
  SetPixelV Form1.hdc, sng1, n2, cL
  Next n2
  Next sng1
 End If
 For ns1 = 1 To pixarray 'clear coordinate array
 savX(ns1) = 0: savy(ns1) = 0
 svX(ns1) = 0: svY(ns1) = 0
 Next ns1
 Form_Paint
 Elap = GetTickCount
 newbackground = False 'we are done with this section until a form resize
End If


'Here is the loop that animates
 'I do not actually use animation 'frames'; rather,
 'I erase an old pixel then draw a new, then complete
 'any missed pixels which usually occur due to the
 'different number of pixels generated each cycle

Do While Not Fin  'Pressing Esc for example..
 DoEvents: If Fin Or breakloop Then Exit Do

 'This part allows for constant rotation speed
 'regardless of cpu power or model complexity.
 'The multiplier (.05) scales the rotation speed
 cL = GetTickCount: sngAP = (cL - Elap) * 0.05
 Elap = cL
 
 ay = ay + ayi * sngAP 'incrementing rotation
 ax = ax + axi * sngAP
 az = az + 0
 If ay > twopi Then
 ay = ay - twopi
 ElseIf ay < 0 Then ay = ay + twopi: End If
 If ax > twopi Then
 ax = ax - twopi
 ElseIf ax < 0 Then ax = ax + twopi: End If
 If az > twopi Then
 az = az - twopi
 ElseIf az < 0 Then az = az + twopi: End If
 snay = Sin(ay): csay = Cos(ay) 'precalc some things
 snax = Sin(ax): csax = Cos(ax)
 snaz = Sin(az): csaz = Cos(az)
 
 For q% = 1 To pointcount Step 1 'performing Euler transform to rotate points
 z = z5(q%) * csay - x5(q%) * snay
 X = x5(q%) * csay + z5(q%) * snay
 Y = y5(q%) * csax - z * snax 'Note: ScaleModel writes to the x5 y5 z5 arrays.
 z = z * csax + y5(q%) * snax 'The 'non-destructed' values are in
 X = X * csaz - Y * snaz      'the x3 y3 z3 arrays.
 Y = Y * csaz + X * snaz
 z = radius * z
 vpd = radius * eye / (eye - z) 'vpd = vanishing-point distortion
 X = X * vpd + sw 'horizontal center
 Y = Y * vpd + sh
 px(q%) = X: pY(q%) = Y
 Next q%
 
 If modelselect = 2 Then 'we're viewing the waving landscape.
 'Here is where we want to apply changes to the height of points
 ax1 = ax1 + 0.2: ay1 = ay1 + 0.2
 If ax1 > twopi Then
 ax1 = ax1 - twopi
 ElseIf ax1 < 0 Then ax1 = ax1 + twopi: End If
 If ay1 > twopi Then
 ay1 = ay1 - twopi
 ElseIf ay1 < 0 Then ay1 = ay1 + twopi: End If
 pointcount = 1
  For sngX = -4.5 To 4.5 Step 1
  For sngY = -4.5 To 4.5 Step 1
  z5(pointcount) = 0.1 * Sin(sngY + sngX + ax1 + ay1)
  pointcount = pointcount + 1
  Next sngY
  Next sngX
 End If

 'Calling antialias( LineNumber ) to generate x, y,
 'alpha values, and the number of pixels to be drawn
 pixarray& = 0
 For q% = 1 To linecount%
 Call antialias(q%)
 Next q%
 
 realpix = 0
 
 'This tricky sub keeps pixels from being drawn twice in same frame
 For ns1& = 1 To pixarray Step 1
 If Not csn(ns1) Then
 x4& = savX&(ns1&): y4& = savy&(ns1&)
 If x4& > -1& And y4& > -1& And x4& < fw& And y4& < fh& Then
 draw(ns1&) = 1
 Select Case bcleansurf(x4&, y4&)
 Case False
 realpix& = realpix& + 1
 bcleansurf(x4&, y4&) = True
 bclneapix(ns1&) = True
 savBG(realpix&) = backbuff&(x4&, y4&)
 Case Else
 bclneapix(ns1&) = False
 End Select
 Else
 draw(ns1&) = 0: End If
 End If
 Next ns1&: pixeln = 0

 cswap = 0 'cswap is used to determine which line we are drawing
 'This heavy loop computes color over background based upon alpha that antialias() generates
 For ns1& = 1 To pixarray& Step 1
 x4& = savX&(ns1&): y4& = savy&(ns1&)
 Select Case csn(ns1&) 'at start of each new line, Sub antialias()
                       'sets csn(first pixel # of that line) = True
                       'For purpose of establishing new line color,
 Case True             'intensity and drawstyle
 cswap = cswap + 1
 If cswap > linecount Then cswap = 0
 intensitybyt = opac(cswap): drawselect = ds(cswap)
 r = cR(cswap): g = cG(cswap): b = cB(cswap)
 'The cR,cG,cB,opac,ds arrays are filled in Sub LineSpec
 csn(ns1&) = 0 'reset
 Case False
 Select Case draw(ns1&)
 Case True
 a! = savAlpha!(ns1&)
 BGR& = backbuff&(x4&, y4&)
 b2& = ((BGR& And &HFF0000) / &H10000) And &HFF
 g2& = ((BGR& And &HFF00) / &H100) And &HFF
 r2% = BGR& And &HFF
 s2! = (intensitybyt / 255) * a!
 Select Case drawselect
 Case 0: cL& = RGB(r2% - s2! * (r2% - r), g2& - s2! * (g2& - g), b2& - s2! * (b2& - b))
 Case 1: cL& = RGB(r2% - s2! * (r2% - a! * r), g2& - s2! * (g2& - a! * g), b2& - s2! * (b2& - a! * b))
 Case 2: cL& = RGB(r2% + s2! * Abs(r2% - r) ^ 0.85, g2& + s2! * Abs(g2& - g) ^ 0.85, b2& + s2! * Abs(b2& - b) ^ 0.85)
 End Select
  Select Case bclneapix(ns1&)
  Case True
  pixeln& = pixeln& + 1&
  savX&(pixeln&) = x4&: savy&(pixeln&) = y4&
  savcL&(pixeln&) = cL&
  backpxln&(x4&, y4&) = pixeln&
  Case Else
  savcL&(backpxln&(x4&, y4&)) = cL&
  End Select
  backbuff&(x4&, y4&) = cL&
  bcleansurf(x4&, y4&) = False
 End Select
 End Select
 Next ns1&

 If arraylen < realpix Then 'If draw pixel# > erase pixel#
  For ns1 = 1 To arraylen Step 1
  x4 = svX(ns1): y4 = svY(ns1): cL = svBG(ns1)
  If cL = backbuff(x4, y4) Then SetPixelV Form1.hdc, x4, y4, cL
  SetPixelV Form1.hdc, savX(ns1), savy(ns1), savcL(ns1)
  Next ns1
  For cL = ns1 To realpix Step 1
  SetPixelV Form1.hdc, savX(cL), savy(cL), savcL(cL)
  Next cL
 Else
  ns1 = realpix + 1
  For cL = ns1 To arraylen Step 1
  x4 = svX(cL): y4 = svY(cL): BGR = svBG(cL)
  If BGR = backbuff(x4, y4) Then SetPixelV Form1.hdc, x4, y4, BGR
  Next cL
  For ns1 = 1 To realpix Step 1
  x4 = svX(ns1): y4 = svY(ns1): cL = svBG(ns1)
  If cL = backbuff(x4, y4) Then SetPixelV Form1.hdc, x4, y4, cL
  SetPixelV Form1.hdc, savX(ns1), savy(ns1), savcL(ns1)
  Next ns1
 End If
 
 'clean backbuffer, store erase pixel information for next cycle
 For ns1 = 1 To realpix Step 1
 cL = savX(ns1): n2 = savy(ns1)
 backbuff(cL, n2) = savBG(ns1)
 svX(ns1) = cL: svY(ns1) = n2
 svBG(ns1) = savBG(ns1)
 Next ns1
 
 'Store this cycle's # of pixels drawn, used in next cycle to erase
 arraylen = realpix

Loop
 
If breakloop Then breakloop = 0: AnimWireframe
Unload Me

End Sub
Private Sub PointDEF(Optional ptX As Single, Optional ptY As Single, Optional ptZ As Single)
pointcount = pointcount + 1
x3(pointcount) = ptX: y3(pointcount) = ptY: z3(pointcount) = ptZ
End Sub
Private Sub LineDEF(Optional LPoint1 As Long, Optional LPoint2 As Long, Optional SameFlavor As Boolean, Optional LnRed As Byte, Optional LnGrn As Byte, Optional LnBlu As Byte, Optional LnIntensity As Byte, Optional LnDrawstyle As Byte)
Dim LCn As Integer
 linecount = linecount + 1
 LCn = linecount
 LnPt1(LCn) = LPoint1: LnPt2(LCn) = LPoint2
 If SameFlavor And LCn > 1 Then
 cR(LCn) = cR(LCn - 1)
 cG(LCn) = cG(LCn - 1)
 cB(LCn) = cB(LCn - 1)
 opac(LCn) = opac(LCn - 1)
 ds(LCn) = ds(LCn - 1)
 Else
 cR(LCn) = LnRed
 cG(LCn) = LnGrn
 cB(LCn) = LnBlu
 opac(LCn) = LnIntensity
 End If
 
 If LnIntensity = 0 And Not SameFlavor Then opac(LCn) = 255
 If LnDrawstyle < 3 And Not SameFlavor Then ds(LCn) = LnDrawstyle
End Sub
Private Sub ScaleModel()
Dim maxlength As Single
'For lngFN = 1 To linecount Step 1
'intN1 = LnPt1(lngFN): intN2 = LnPt2(lngFN)
'sngAP = Sqr((x3(intN1) - x3(intN2)) ^ 2 + _
            (y3(intN1) - y3(intN2)) ^ 2 + _
            (z3(intN1) - z3(intN2)) ^ 2)
For lngFN = 1 To pointcount Step 1
sngAP = Sqr(x3(lngFN) ^ 2 + y3(lngFN) ^ 2 + z3(lngFN) ^ 2)
If sngAP > maxlength Then maxlength = sngAP
Next lngFN

For lngFN = 1 To pointcount Step 1
x5(lngFN) = 0.78 * x3(lngFN) / maxlength
y5(lngFN) = 0.78 * y3(lngFN) / maxlength
z5(lngFN) = 0.78 * z3(lngFN) / maxlength
Next lngFN

End Sub

' E S S E N T I A L - numbers in here not meant to be messed with
Private Sub antialias(LineN As Integer)
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
Dim spx As Double, epx As Double
Dim spy As Double, epy As Double
Dim ax As Double, bx As Double, cx As Double, dx As Double
Dim ay As Double, by As Double, cy As Double, dy As Double
Dim ex As Double, ey As Double
Dim mp5 As Single, pp5 As Single
Dim rex As Integer, rey As Integer

Dim trz As Double, tri As Double
Dim lwris As Double, lwrun As Double
Dim zsl As Single
Dim slope As Double, lsope As Double
Dim midx As Double, midy As Double
Dim sl2 As Single
Dim distanc1 As Double, distanc2 As Double
Dim diagonal As Boolean
Dim a As Single
Dim st As Integer, lp1 As Integer, lp2 As Integer
Dim one As Single

pixarray = pixarray + 1  'Here, we are faking +1 to pixel count before
           'any pixel data on a new line is computed.
           
           'This value will be used in a loop within
           'AnimWireframe to determine at which pixel
           'number does a new line start.

csn(pixarray) = 1 'So, at this n in the array, AnimWireframe will
           'know when to change r,g,b and intensitybyt
           'values based upon specific line 'properties'
           'that are assigned to cR,cG,cB,opac arrays.
           
           'To see where this happens, Find " Case csn("

'Okay, time to get computing.
lp1 = LnPt1(LineN): lp2 = LnPt2(LineN)
x1 = px(lp1): y1 = pY(lp1)
x2 = px(lp2): y2 = pY(lp2)

If x1 < x2 Then
 epy = y2: epx = x2
 spy = y1: spx = x1
Else
spy = y2: spx = x2
epy = y1: epx = x1: End If

If epx = spx Or epy = spy Then
 diagonal = 0
 If epy > spy Then
  st = 1
 Else: st = -1: End If

Else: slope = (epy - spy) / (epx - spx): lsope = -1 / slope
diagonal = 1: End If

midx = 0.5 * lsope: midy = 0.5 * slope
sl2 = slope * slope: one = 1 / Sqr(1 + sl2)
lwris = 1 - (one - Abs(slope) + sl2 * one)
lwrun = lwris * Abs(lsope)
distanc1 = 0.5 * one
distanc2 = slope * distanc1
ax = spx - distanc1 - distanc2
ay = spy + distanc1 - distanc2
bx = epx + distanc1 - distanc2
by = epy + distanc1 + distanc2
cx = epx + distanc1 + distanc2
cy = epy - distanc1 + distanc2
dx = spx - distanc1 + distanc2
dy = spy - distanc1 - distanc2

one = 255 * (1 - 0.5 * lwris * lwrun)

If diagonal Then
If slope > 0 Then
If slope <= 1 Then
ey# = slope# * (Round(ax#) + 1.5 - ax#) + ay#
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1
tri# = pp5! - lwris#: trz# = pp5! - slope#: zsl! = mp5! - midy
For ex# = Round(ax#) + 1.5 To Round(bx#) - 1.5 Step 1
pixarray = pixarray + 1
 If ey# > tri# Then
  savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(one!)
 Else
  If ey# > trz# Then
  savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (1 + lsope# * sr(Int((pp5! - ey#) * 128))))
  Else: savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (ey# - zsl!)): End If
 End If
If ey# > trz# Then
 pp5! = pp5! + 1: tri# = tri# + 1: trz# = trz# + 1
 zsl! = zsl! + 1: rey% = rey% + 1: End If
ey = ey# + slope: Next ex#

ey# = cy# - slope# * (cx# - Round(cx#) + 1.5)
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = pp5! + midy
tri# = mp5! + lwris: trz# = mp5! + slope
For ex# = Round(cx#) - 1.5 To Round(dx#) + 1.5 Step -1
If ey# > tri# Then
pixarray = pixarray + 1
 If ey# > trz# Then
 savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (zsl! - ey#))
 Else: savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (1 + lsope# * sr(Int((ey# - mp5!) * 128))))
 End If
End If
If ey# < trz# Then
 mp5! = mp5! - 1: tri# = tri# - 1: trz# = trz# - 1
 zsl! = zsl! - 1: rey% = rey% - 1: End If
ey# = ey# - slope: Next ex#

ex# = cx# + lsope * (cy# - Round(cy#) + 0.5)
For ey# = Round(cy#) - 0.5 To Round(dy#) + 1.5 Step -1
pixarray = pixarray + 1
 rex% = Round(ex#)
 savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (slope# * sr(Int((ex# - rex% + 0.5) * 128))))
ex# = ex# + lsope#: Next ey#

ex# = ax# - lsope# * (Round(ay#) + 0.5 - ay#)
For ey# = Round(ay) + 0.5 To Round(by) - 0.5 Step 1
pixarray = pixarray + 1
 rex% = Round(ex#)
 savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (slope# * sr(Int((rex% + 0.5 - ex#) * 128))))
ex# = ex# - lsope#: Next ey#

Else
ex# = dx# - lsope# * (Round(dy#) + 1.5 - dy#): rex% = Round(ex#): mp5! = rex% - 0.5
pp5! = mp5! + 1: tri# = pp5! - lwrun: trz# = pp5! + lsope: zsl! = mp5! + midx
For ey# = Round(dy#) + 1.5 To Round(cy#) - 0.5 Step 1
pixarray = pixarray + 1
 If ex# > tri# Then
  savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(one!)
 Else
  If ex# > trz# Then
   savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (1 - slope * sr(Int((pp5! - ex#) * 128))))
  Else: savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (ex# - zsl!)): End If
 End If
If ex# > trz# Then
 tri# = tri# + 1: trz# = trz# + 1: pp5! = pp5! + 1
 zsl! = zsl! + 1: rex% = rex% + 1: End If
ex# = ex# - lsope#: Next ey#

ex# = bx# + lsope# * (by# - Round(by#) + 0.5): rex% = Round(ex#): mp5! = rex% - 0.5
pp5! = mp5! + 1: tri# = mp5! + lwrun: trz# = mp5! - lsope: zsl! = pp5! - midx#
For ey# = Round(by#) - 0.5 To Round(ay#) + 1.5 Step -1
If ex# > tri# Then
 pixarray = pixarray + 1
 If ex# < trz# Then
  savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (1 - slope# * sr(Int((ex# - mp5!) * 128))))
 Else: savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (zsl! - ex#)): End If
End If
If ex# < trz# Then
 tri# = tri# - 1: trz# = trz# - 1: mp5! = mp5! - 1
 rex% = rex% - 1: zsl! = zsl! - 1: End If
ex# = ex# + lsope#: Next ey#

ey# = dy# + slope# * (Round(dx#) + 0.5 - dx#)
For ex# = Round(dx#) + 0.5 To Round(cx#) - 1.5 Step 1
 pixarray = pixarray + 1
 rey% = Round(ey#)
 savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (-lsope# * sr(Int((rey% + 0.5 - ey#) * 128))))
ey# = ey# + slope#: Next ex#

ey# = ay# + slope# * (Round(ax#) + 1.5 - ax#)
For ex# = Round(ax#) + 1.5 To Round(bx#) - 0.5 Step 1
 pixarray = pixarray + 1
 rey% = Round(ey#)
 savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (-lsope# * sr(Int((ey# - rey% + 0.5) * 128))))
ey# = ey# + slope: Next ex#
End If

Else
If slope > -1 Then
ey# = dy# + slope# * (Round(dx#) + 1.5 - dx#)
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = pp5! - midy#
tri# = mp5! + lwris: trz# = mp5! - slope#
For ex# = Round(dx#) + 1.5 To Round(cx#) - 1.5 Step 1
pixarray = pixarray + 1
 If ey# < tri# Then
  savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(one!)
 Else
  If ey# < trz# Then
   savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (1 - lsope# * sr(Int((ey# - mp5!) * 128))))
  Else: savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (zsl! - ey#)): End If
 End If
If ey# < trz# Then
  mp5! = mp5! - 1: zsl! = zsl! - 1: tri# = tri# - 1
  rey% = rey% - 1: trz# = trz# - 1: End If
ey# = ey# + slope#: Next ex#

ey# = by# - slope# * (bx# - Round(bx#) + 1.5)
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = mp5! + midy#
tri# = pp5! - lwris: trz# = pp5! + slope#
For ex# = Round(bx#) - 1.5 To Round(ax#) + 1.5 Step -1
If ey# < tri# Then
 pixarray = pixarray + 1
 If ey# < trz# Then
  savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (ey# - zsl!))
 Else
  savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (1 - lsope# * sr(Int((pp5! - ey#) * 128))))
 End If
End If
If ey# > trz# Then
  rey% = rey% + 1: pp5! = pp5! + 1: zsl! = zsl! + 1
  tri# = tri# + 1: trz# = trz# + 1: End If
ey# = ey# - slope#: Next ex#

ex# = ax# + lsope# * (ay# - Round(ay#) + 1.5)
For ey# = Round(ay#) - 1.5 To Round(by#) + 0.5 Step -1
pixarray = pixarray + 1
 rex% = Round(ex#)
 savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (-slope * sr(Int((ex# - rex% + 0.5) * 128))))
ex# = ex# + 2 * midx#: Next ey#

ex# = cx# - lsope# * (Round(cy#) + 1.5 - cy#)
For ey# = Round(cy#) + 1.5 To Round(dy#) - 0.5
pixarray = pixarray + 1
 rex% = Round(ex#)
 savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (-slope# * sr(Int((rex% + 0.5 - ex#) * 128))))
ex# = ex# - 2 * midx#: Next ey#

Else
ex# = ax# + lsope# * (ay# - Round(ay#) + 1.5): rex% = Round(ex#): mp5! = rex% - 0.5
zsl! = mp5! - midx#: pp5! = mp5! + 1: tri# = pp5! - lwrun#: trz# = pp5! - lsope#
For ey# = Round(ay#) - 1.5 To Round(by#) + 1.5 Step -1
pixarray = pixarray + 1
 If ex# > tri# Then
  savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(one!)
 Else
  If ex# > trz# Then
   savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (1 + slope# * sr(Int((pp5! - ex#) * 128))))
  Else: savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (ex# - zsl!)): End If
 End If
If ex# > trz# Then
  tri# = tri# + 1: trz# = trz# + 1: pp5! = pp5! + 1
  rex% = rex% + 1: zsl! = zsl! + 1: End If
ex# = ex# + lsope#: Next ey#

ex# = cx# - lsope# * (Round(cy#) + 1.5 - cy#): rex% = Round(ex#): mp5! = rex% - 0.5
tri# = mp5! + lwrun#: trz# = mp5! + lsope#: zsl! = mp5! + 1 + midx#
For ey# = Round(cy#) + 1.5 To Round(dy#) - 1.5 Step 1
 If ex# > tri# Then
 pixarray = pixarray + 1
  If ex# < trz# Then
   savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (1 + slope * sr(Int((ex# - mp5!) * 128))))
  Else: savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (zsl! - ex#)): End If
 End If
If ex# < trz# Then
 tri# = tri# - 1: trz# = trz# - 1: zsl! = zsl! - 1
 rex% = rex% - 1: mp5! = mp5! - 1: End If
ex# = ex# - lsope#: Next ey#

ey# = by# - slope# * (bx# - Round(bx#) + 2.5)
For ex# = Round(bx#) - 2.5 To Round(ax#) + 0.5 Step -1
pixarray = pixarray + 1
 rey% = Round(ey#)
 savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (lsope# * sr(Int((ey# - rey% + 0.5) * 128))))
ey# = ey# - slope#: Next ex#

ey# = dy# + slope# * (Round(dx#) + 1.5 - dx#)
For ex# = Round(dx#) + 1.5 To Round(cx#) - 0.5 Step 1
pixarray = pixarray + 1
 rey% = Round(ey#)
 savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (lsope# * sr(Int((rey% + 0.5 - ey#) * 128))))
ey# = ey# + slope#: Next ex#

End If
End If

Else
If epy = spy Then
 rey% = Round(ay): st% = rey% - 1: a! = ay - rey% + 0.5
For ex# = Round(ax) + 1.5 To Round(bx) - 0.5
 pixarray = pixarray + 1
 zsl! = gs(255! * a!): rex% = Int(ex#)
 savX(pixarray) = rex%: savy(pixarray) = rey%: savAlpha(pixarray) = zsl!
 pixarray = pixarray + 1
 savX(pixarray) = rex%: savy(pixarray) = st%: savAlpha(pixarray) = 1 - zsl!
Next ex#
ElseIf epx = spx Then
 rex% = Round(ax): st% = rex% - 1: a! = ax - rex% + 0.5
 If epy > spy Then
 trz = 1
 Else: trz = -1: End If
For ey# = Round(ay) - 0.5 To Round(by) + 1.5 Step trz
 pixarray = pixarray + 1
 zsl! = gs(255! * a!): rey% = Int(ey#)
 savX(pixarray) = rex%: savy(pixarray) = rey%: savAlpha(pixarray) = zsl!
 pixarray = pixarray + 1
 savX(pixarray) = st%: savy(pixarray) = rey%: savAlpha(pixarray) = 1 - zsl!
Next ey#: End If
End If
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case 48 To 57 'number keys
modelselect = KeyCode - 49
LineSpec
 
Case 105, 73 'i
If bwi Then
bwi = 0
Else
bwi = 1
End If
newbackground = True: breakloop = True
Case Else: Fin = 1
End Select

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 selectv = 0
 pressed = 1
 For bytW = 1 To 1
 Select Case X
 Case vleft(bytW) To vright(bytW)
 Select Case Y
 Case vtop(bytW) To vbot(bytW)
 selectv = bytW 'We've landed inside a control's dimensions
 yInit = 0
  End Select
   End Select
    Next bytW
 
 Select Case selectv
 Case 0
  xr = X
  yr = Y
  axi = axi * 0.7
  ayi = ayi * 0.7
 End Select

 Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pressed Then
If selectv > 0 Then
Select Case Y
Case Is <> yInit
yInit = Y: bytW = selectv
 Select Case Y
 Case vtop(bytW) To vbot(bytW)
 vval(bytW) = vmax(bytW) * (vbot(bytW) - Y) / (vbot(bytW) - vtop(bytW))
 Case Is < vtop(bytW)
 vval(bytW) = vmax(bytW)
 Case Is > vbot(bytW)
 vval(bytW) = vmin(bytW)
 End Select
 
 Select Case bytW
 Case 1
 shape = vval(bytW)
 pow = 1 / (1 + shape / 255)
 gs(0) = 0: gs(255) = 1
 For bytW = 1 To 254
 gs(bytW) = (bytW / 255) ^ pow
 Next bytW
 
 End Select
End Select
Else
 xr = xr - X
 If xr > 0 Then
  If xr > xr2 Then
   xr2 = xr
  End If
  ayi = ayi * 0.8
  If xr > 6 Then
   ayi = ayi + xr2 / 1000
  Else
   ayi = ayi + xr / 500: xr2 = xr2 - xr
  End If
 
 ElseIf xr < 0 Then
  If xr < xr2 Then
   xr2 = xr
  End If
  ayi = ayi * 0.8
  If xr < -6 Then
   ayi = ayi + xr2 / 1000
  Else
   ayi = ayi + xr / 500: xr2 = xr2 - xr
  End If
 End If
 
 yr = yr - Y
 If yr > 0 Then
  If yr > yr2 Then
   yr2 = yr
  End If
  axi = axi * 0.8
  If yr > 6 Then
   axi = axi + yr2 / 1000
  Else
   axi = axi + yr / 500: yr2 = yr2 - yr
  End If
 
 ElseIf yr < 0 Then
  If yr < yr2 Then
   yr2 = yr
  End If
  axi = axi * 0.8
  If yr < -6 Then
   axi = axi + yr2 / 1000
  Else
   axi = axi + yr / 500: yr2 = yr2 - yr
  End If
 End If
 
 xr = X
 yr = Y

End If
End If
End Sub
Private Sub Form_DblClick()
pressed = 1
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pressed = 0
Select Case selectv
Case 1

End Select
End Sub
Private Sub Form_Paint()
Static outline As Long, outline2 As Long
 If Not wildbackground Then
 'screen colorfade effect
 Select Case bwi
 Case True
 Form1.BackColor = vbBlack
 Case Else
 Form1.BackColor = vbWhite
 End Select
 For h = 255 To 0 Step -1
 q = 741 - h * 5
 dr = ditr / 255 * h
 dg = ditg / 255 * h
 db = ditb / 255 * h
 Select Case bwi
 Case 0: rr = 255 - dr: gg = 255 - dg: bb = 255 - db: q = 741 - h * 3
 Case Else: rr = dr: gg = dg: bb = db: q = 601 - h * 3
 End Select
 BGR = RGB(rr, gg, bb)
 Form1.Line (0, q)-(fw, q), BGR, B
 Form1.Line (0, q - 1)-(fw, q - 1), BGR, B
 Form1.Line (0, q - 2)-(fw, q - 2), BGR, B
 Next h
 End If
 
 If bwi Then
  Form1.ForeColor = RGB(100, 75, 190)
  outline = RGB(70, 86, 90)
  outline2 = RGB(254, 78, 81)
 Else
  Form1.ForeColor = RGB(185, 140, 255)
  outline = RGB(140, 70, 216)
  outline2 = RGB(0, 0, 200)
  Form1.ForeColor = RGB(0, 195, 198)
  outline = RGB(50, 0, 196)
 End If

 'outlines of buttons and sliders
 For bytW = 1 To 1 Step 1
  Select Case bytW
  Case 1
  Form1.Line (vleft(bytW), vtop(bytW))-(vright(bytW), vbot(bytW)), outline2, B
  Form1.Print "Press " & Chr(34) & "i" & Chr(34) & " to change background"
  End Select
 Next bytW
   
End Sub
Private Sub Form_Resize()
fw = Form1.ScaleWidth: fw1 = fw
fh = Form1.ScaleHeight: fw2 = fh
sw = fw / 2
sh = fh / 2

vleft(1) = 10: vright(1) = vleft(1) + 8
vtop(1) = 10: vbot(1) = vtop(1) + 100
vmax(1) = 255

If fw > 0 Then eye = 1.6 * sw

radius = sw / 2: breakloop = True: newbackground = True
End Sub
Private Sub Form_Load()
Randomize
Form1.ScaleMode = vbPixels

wildbackground = True 'bok bok bok bok b'gah
bwi = 0    'Boolean - 0 for light background
           'bwi is changed when you type 'i'

'This array allows me to remove a multiply in a few spots
'in antialias()
sr(0) = 0: sr(128) = 0.5
For bytW = 1 To 127
sr(bytW) = (bytW / 128) ^ 2 / 2
 Next bytW

shape = 10 'byte used to apply a curve to 'brightness'
           'values in gs() array - used by antialias()
           
           'Purpose is to adjust the look of the antialias effect
           'Usually a higher setting will look better over a dark
           'background

pow = 1 / (1 + shape / 255)
gs(0) = 0: gs(255) = 1
For bytW = 1 To 254
gs(bytW) = (bytW / 255) ^ pow
 Next bytW
'This array is also adjusted in Form_MouseMove

'Random starting rotation speed
axi = 0.01 * (Rnd - 0.5): ayi = 0.02 * (Rnd - 0.5)
'axi = 0.001: ayi = 0.04: modelselect = 1
ax = pi

End Sub
Private Sub Form_Unload(Cancel As Integer)
Fin = True
End Sub
