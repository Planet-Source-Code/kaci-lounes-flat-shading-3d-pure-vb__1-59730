VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":0000
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##########################################################
'#  Author: Mr KACI Lounes                                #
'#  A 3D FlatShading Engine in *Pure* VB Code!            #
'#  Compile for more speed                                #
'#  Mail me at KLKEANO@HOTMAIL.COM                        #
'#  Copyright © 2005 - KACI Lounes - All rights reserved. #
'##########################################################

Option Explicit                  'Stop undeclared variables

Const FOV% = 400                 'Field of view, or the distance between
                                 'the eye and the projection plane

Const NearPlane! = 0             'ClipFar & ClipNear Zs values
Const FarPlane! = 500            '       ,,           ,,

Dim Scene() As Mesh              'Array contains the data structures for all objects
Dim Spots(1) As SpotLight3D      'For lighting the scene, we use two Spotlights.

Dim FacesDepth() As Single       'Sort Arrays: Depths array
Dim FacesIndex() As Long         '             Faces index array
Dim MeshsIndex() As Byte         '             Mesh index array

Dim CamPos As Vector3D           'Camera vector
Dim CamDir As Vector3D           'Camera Direction (Pitch/Yaw mode)
Dim CamTo As Vector3D            'Camera LookAt (Pitch/Yaw mode)
Dim ViewMatrix As Matrix         'View matrix
Dim ViewMode As Boolean          'View mode, choose between LookAt(True), or Pitch/Yaw(False)
Dim FogEnable As Boolean         'Fogging mode
Dim LookAtObj%                   'Which object to look at ? (1, 2, 3)
Dim XAng%, YAng%                 'Anlges rotations (Pitch/Yaw mode)
Dim Ax!, Ay!                     '         ,,             ,,

Dim TeapAng%                     'Teapot ZAngle rotation
Dim I&, J&                       'Loops counters

Dim Pts() As Point2D             'Rasterization vars
Dim Vs() As Point2D
Dim Ts(50) As Triangle
Dim Tri(2) As Point2D
Dim Stat As Byte
Dim NumPts As Byte
Dim NumTris As Integer
Dim FInx&, MInx%, S%
Dim X1!, Y1!, X2!, Y2!, X3!, Y3!

                                 'Booleans for processing keyboard input:

Dim KeyESC As Boolean            'Escape Key     : Exit program
Dim KeySPC As Boolean            'Space Key      : Change view mode
Dim KeyHom As Boolean            'Home Key       : Change X Angle (+)
Dim KeyEnd As Boolean            'End Key        : Change X Angle (-)
Dim KeyLft As Boolean            'Left Key       : Change Y Angle (+)
Dim KeyRgt As Boolean            'Right Key      : Change Y Angle (-)
Dim KeyTop As Boolean            'Up Key         : Move front (Z+)
Dim KeyBot As Boolean            'Down Key       : Move back (Z-)
Dim KeyF As Boolean              'F Key:         : Enable/Disable Fogging mode
Dim KeyPad1 As Boolean           'NumPad 1 Key   : Look at the Sphere (LookAt mode only)
Dim KeyPad2 As Boolean           'NumPad 2 Key   : Look at the Teapot (LookAt mode only)
Dim KeyPad3 As Boolean           'NumPad 3 Key   : Look at the Torus (LookAt mode only)
Dim KeyPad4 As Boolean           'NumPad 4 Key   : Enable/Disable Spot1
Dim KeyPad5 As Boolean           'NumPad 5 Key   : Enable/Disable Spot2
Sub Shade(MeshIndex As Long, FaceIndex As Long)

 'Shade faces:
 '============
 'Important Note: For Gouraud Shading, we use the FaceCenter as V1, V2, V3.
 '                Then, we obtains tree colors, so we lineary interpolate
 '                 these colors for 2D rasterization (Gradient Fill).
 '                Note also for the Phong shade, we use ALL pixels as vectors
 '                 and shade them (slow).

 Dim ColorSum As ColRGB, I%
 Dim Alpha!, Beta!, Gamma!, Delta!, Epsilon!   ' 'Single' values, usally: (0....1)
 '  Diffuse  Spec   Attenu   Fog     Spot

 For I = 0 To 1                       'For each spot,
  If Spots(I).Enabled = True Then     ' and only if the spot is turned on:

   '==================================================
   'Note: For the PointLight algorithm (Sphere of light) or 'Omni',
   '       you change simply the Epsilon value, in others words,
   '        Epsilon is responsible for the shape of light.

   'Use SpotLight filter:
   Epsilon = 1
   If Spots(I).Falloff > 0 Then
    Epsilon = VectorAngle(Spots(I).TDir, VectorSubtract3(Scene(MeshIndex).Faces(FaceIndex).Center, Spots(I).TPos))
    If Epsilon < 0 Then Epsilon = 0
    If Spots(I).Falloff <> Spots(I).Hotspot Then
     Epsilon = (Spots(I).Falloff - Epsilon) / (Spots(I).Falloff - Spots(I).Hotspot)
     If Epsilon < 0 Then Epsilon = 0 Else If Epsilon > 1 Then Epsilon = 1
    Else
     Exit Sub
    End If
   End If

   '==================================================
   'Perform incident ray shading (Diffusion):
   Alpha = VectorAngle(VectorSubtract3(Spots(I).TPos, Scene(MeshIndex).Faces(FaceIndex).Center), Scene(MeshIndex).Faces(FaceIndex).Normal)
   Alpha = (Alpha * Spots(I).Diffusion)
   If Alpha < 0 Then Alpha = 0

   '==================================================
   'Perform reflected ray shading (Specular):
   Beta = VectorAngle(VectorReflect(VectorSubtract3(Scene(MeshIndex).Faces(FaceIndex).Center, Spots(I).TPos), Scene(MeshIndex).Faces(FaceIndex).Normal), VectorSubtract3(CamPos, Scene(MeshIndex).Faces(FaceIndex).Center))
   Beta = (Beta * Spots(I).Specular)
   If Beta < 0 Then Beta = 0

   '==================================================
   'Apply light distance decay (Attenuation):
   Gamma = 1
   If Spots(I).AttenEnable = True Then
    If Spots(I).DarkRange <> Spots(I).BrightRange Then
     Gamma = (Spots(I).DarkRange - VectorDistance3(Scene(MeshIndex).Faces(FaceIndex).Center, Spots(I).TPos)) / (Spots(I).DarkRange - Spots(I).BrightRange)
     If Gamma < 0 Then Gamma = 0 Else If Gamma > 1 Then Gamma = 1
    Else
     Exit Sub      'Bads parameters !
    End If
   End If

   '==================================================
   'Here, we have four values: Epsilon, Alpha, Beta & Gamma
   ' The next step is 'Averaging' these values to obtains the face color,
   '  This is known by: 'The Lighting Model'.
   '
   'In this application, The lighting model is: (Alpha + Beta) * Gamma * Epsilon
   'Note that always: (Alpha + Beta), in other word: (Diffusion + Specular)
   '
   ' In the next line, we 'Add' the color to apply 'Multiples Lights':

   ColorSum = ColorAdd(ColorSum, ColorScale(Spots(I).Color, ((Alpha + Beta) * Gamma) * Epsilon))
   ColorSum = ColorInterpolate(ColorSum, Scene(MeshIndex).Faces(FaceIndex).Col, Scene(MeshIndex).Faces(FaceIndex).ShadVal)
   ColorSum = ColorAdd(ColorSum, Spots(I).Ambiance)  'Add the Ambiance

  End If
 Next I

 'Set limitations:
 If ColorSum.R > 255 Then ColorSum.R = 255 Else If ColorSum.R < 0 Then ColorSum.R = 0
 If ColorSum.G > 255 Then ColorSum.G = 255 Else If ColorSum.G < 0 Then ColorSum.G = 0
 If ColorSum.B > 255 Then ColorSum.B = 255 Else If ColorSum.B < 0 Then ColorSum.B = 0

 'Here, Lighting is Ok, the follow operation
 'is simply 'Scale' the color for apply the Fog,
 'with the NearPlane and FarPlane properties of the camera,
 ' and, very sure, if only the Fog is enable.

 '==================================================
 'Apply camera distance decay (Fog)

 Delta = 1
 If FogEnable = True Then
  Delta = (FarPlane - VectorDistance3(Scene(MeshIndex).Faces(FaceIndex).Center, CamPos)) / (FarPlane - NearPlane)
  If Delta < 0 Then Delta = 0 Else If Delta > 1 Then Delta = 1
 End If

 '==================================================
 Scene(MeshIndex).Faces(FaceIndex).TCol = ColorScale(ColorSum, Delta)

End Sub
Sub Process()

 'Calculate the view matrix:
 '==========================
 If ViewMode = True Then           'LookAt mode at LookAtObj%:
  Select Case LookAtObj
   Case 1: ViewMatrix = MatrixView(CamPos, VectorInput3(30, -10, 0), VectorInput3(0, 1, 0))   'At Sphere
   Case 2: ViewMatrix = MatrixView(CamPos, VectorInput3(0, 0, 30), VectorInput3(0, 1, 0))     'At Teapot
   Case 3: ViewMatrix = MatrixView(CamPos, VectorInput3(-20, -5, -10), VectorInput3(0, 1, 0)) 'At Torus
  End Select
 Else                              'Pitch/Yaw mode:
  CamDir.X = Sin(XAng * Deg) * Cos(YAng * Deg)
  CamDir.Y = Sin(YAng * Deg)
  CamDir.Z = Cos(XAng * Deg) * Cos(YAng * Deg)
  CamTo = VectorAdd3(CamDir, CamPos)
  ViewMatrix = MatrixView(CamPos, CamTo, VectorInput3(0, 1, 0))
 End If

 'Prepare teapot matrix for ZRotation (Roll):
 '===========================================
 Scene(2).IDMatrix = MatrixWorld(VectorInput3(0, -5, 30), VectorInput3(0.5, 0.5, 0.5), Deg * 90, (TeapAng * Deg), 0)

 'Transorm & Project :
 '====================
 For J = LBound(Scene) To UBound(Scene)
  For I = LBound(Scene(J).Vertices) To UBound(Scene(J).Vertices)

   'World transformation:
   Scene(J).TmpVerts(I) = MatrixMultiplyVector3(Scene(J).Vertices(I), Scene(J).IDMatrix)
   'View transformation:
   Scene(J).TmpVerts(I) = MatrixMultiplyVector3(Scene(J).TmpVerts(I), ViewMatrix)

   'Projection (Persective Distortion): (you can change the FOV):
   'Ignore the division by zero, by remplacing it by 0.0001:
   If Scene(J).TmpVerts(I).Z = 0 Then Scene(J).TmpVerts(I).Z = 0.0001
   '
   'Apply the persective distortion ((X/Z) * FOV, (Y/Z) * FOV):
   'For an orthographic projection, We simply skip the next two lines:
   Scene(J).TmpVerts(I).X = (Scene(J).TmpVerts(I).X / Scene(J).TmpVerts(I).Z) * FOV
   Scene(J).TmpVerts(I).Y = (Scene(J).TmpVerts(I).Y / Scene(J).TmpVerts(I).Z) * FOV

  Next I
 Next J

 'Transform the spots vectors:
 '============================
 For I = LBound(Spots) To UBound(Spots)
  Spots(I).TPos = MatrixMultiplyVector3(Spots(I).Origin, ViewMatrix)
  Spots(I).TDir = MatrixMultiplyVector3(Spots(I).Direction, ViewMatrix)
 Next I

 'Hidden faces removal (i complete this in nexts versions):
 '=========================================================
 ' 1- Check the visiblity of faces by the normal
 ' 2- Check if the triangle is between FarPlane & NearPlane
 '
 For J = LBound(Scene) To UBound(Scene)
  For I = LBound(Scene(J).Faces) To UBound(Scene(J).Faces)
   'Get the face normal :
   Scene(J).Faces(I).Normal = VectorGetNormal(Scene(J).TmpVerts(Scene(J).Faces(I).A), Scene(J).TmpVerts(Scene(J).Faces(I).B), Scene(J).TmpVerts(Scene(J).Faces(I).C))

   If Scene(J).Faces(I).Normal.Z > 0 Then
    '--------------------------------------
    If Scene(J).TmpVerts(Scene(J).Faces(I).A).Z > NearPlane And _
       Scene(J).TmpVerts(Scene(J).Faces(I).B).Z > NearPlane And _
       Scene(J).TmpVerts(Scene(J).Faces(I).C).Z > NearPlane Then
     '-------------------------------------------------------
     If Scene(J).TmpVerts(Scene(J).Faces(I).A).Z < FarPlane And _
        Scene(J).TmpVerts(Scene(J).Faces(I).B).Z < FarPlane And _
        Scene(J).TmpVerts(Scene(J).Faces(I).C).Z < FarPlane Then

      'Face is visible, then:

      'Calculate the face center:
      Scene(J).Faces(I).Center = VectorGetCenter(Scene(J).TmpVerts(Scene(J).Faces(I).A), Scene(J).TmpVerts(Scene(J).Faces(I).B), Scene(J).TmpVerts(Scene(J).Faces(I).C))
      Shade J, I

      'Add the averaged depth of face to FacesDepths array:
      ReDim Preserve FacesDepth(UBound(FacesDepth) + 1)
      FacesDepth(UBound(FacesDepth)) = (Scene(J).TmpVerts(Scene(J).Faces(I).A).Z + _
                                        Scene(J).TmpVerts(Scene(J).Faces(I).B).Z + _
                                        Scene(J).TmpVerts(Scene(J).Faces(I).C).Z) * 0.3333333
      'Add the face index to the FacesIndex array:
      ReDim Preserve FacesIndex(UBound(FacesIndex) + 1)
      FacesIndex(UBound(FacesIndex)) = I

      'Add the mesh index to the MeshsIndexs array:
      ReDim Preserve MeshsIndex(UBound(MeshsIndex) + 1)
      MeshsIndex(UBound(MeshsIndex)) = J

     End If
    End If
   End If

  Next I
 Next J

 ExtractSort3D FacesDepth(), FacesIndex(), MeshsIndex()     'Sort back to front

 If TeapAng = 359 Then TeapAng = 0         'For 'Looping' the teapot rotation
 TeapAng = TeapAng + 1

End Sub
Sub GetKeys()

 'Process keyboard entry:
 '=======================
 If KeyESC = True Then Unload Me: End
 If KeySPC = True Then ViewMode = Not ViewMode: KeySPC = False

 '(D'ont worry, The nexts two lines are for moving the player)
 If (YAng < 90) Or (YAng > 270) Then Ax = Sin(XAng * Deg) Else Ax = -Sin(XAng * Deg)
 If (YAng < 90) Or (YAng > 270) Then Ay = Cos(XAng * Deg) Else Ay = -Cos(XAng * Deg)

 If KeyTop = True Then
  CamPos.Z = (CamPos.Z + Ay)
  CamPos.X = (CamPos.X + Ax)
  KeyTop = False
 End If

 If KeyBot = True Then
  CamPos.Z = (CamPos.Z - Ay)
  CamPos.X = (CamPos.X - Ax)
  KeyBot = False
 End If

 If KeyF = True Then FogEnable = Not FogEnable: KeyF = False

 If KeyHom = True Then YAng = YAng + 3: KeyHom = False
 If KeyEnd = True Then YAng = YAng - 3: KeyEnd = False
 If KeyRgt = True Then XAng = XAng + 3: KeyRgt = False
 If KeyLft = True Then XAng = XAng - 3: KeyLft = False

 If KeyPad1 = True Then LookAtObj = 1: KeyPad1 = False
 If KeyPad2 = True Then LookAtObj = 2: KeyPad2 = False
 If KeyPad3 = True Then LookAtObj = 3: KeyPad3 = False

 If KeyPad4 = True Then Spots(0).Enabled = Not Spots(0).Enabled: KeyPad4 = False
 If KeyPad5 = True Then Spots(1).Enabled = Not Spots(1).Enabled: KeyPad5 = False

End Sub
Sub LoadScene()

 Dim File3DName$, Buff&  'Buff& is a 'Long' data type, so coded into 4 bytes (32 Bits)
                         'then note that the Get function recieve 4 bytes for each
                         ' 'Long' (or Single) data type readed.

 ReDim Scene(3)          'Redim the scene for 4 meshs (Grid, Sphere, Teapot & Torus)

 'Load the models:
 '================
 For J = 0 To UBound(Scene)

  Select Case J
   Case 0: File3DName = App.Path & "\Primatives\Grid.klf"
   Case 1: File3DName = App.Path & "\Primatives\Sphere.klf"
   Case 2: File3DName = App.Path & "\Primatives\Teapot.klf"
   Case 3: File3DName = App.Path & "\Primatives\Torus.klf"
  End Select

  Open File3DName For Binary As 1  'Open the file:
   Get #1, , Buff                  'Number of vertices
   ReDim Scene(J).Vertices(Buff)
   ReDim Scene(J).TmpVerts(Buff)
   Get #1, , Buff                  'Number of faces
   ReDim Scene(J).Faces(Buff)
   For I = LBound(Scene(J).Vertices) To UBound(Scene(J).Vertices) 'Read vertices
    Get #1, , Scene(J).Vertices(I).X
    Get #1, , Scene(J).Vertices(I).Y
    Get #1, , Scene(J).Vertices(I).Z
   Next I
   For I = LBound(Scene(J).Faces) To UBound(Scene(J).Faces)       'Read faces
    Get #1, , Scene(J).Faces(I).A
    Get #1, , Scene(J).Faces(I).B
    Get #1, , Scene(J).Faces(I).C
   Next I
  Close 1                          'Close the file

  'Set the world matrix for each object in scene:
  '==============================================
  Select Case J
   'The grid should be big that others objects (as a floor):
   Case 0: Scene(0).IDMatrix = MatrixWorld(VectorInput3(0, -1, 0), VectorInput3(-1.5, 1.5, 1.5), 0, 0, 0)
   Case 1: Scene(1).IDMatrix = MatrixWorld(VectorInput3(30, -10, 0), VectorInput3(0.3, 0.3, 0.3), 0, 0, 0)
   'If you change the next line (Teapot), Don't forget to update it in the 'Process' routine.
   Case 2: Scene(2).IDMatrix = MatrixWorld(VectorInput3(0, -5, 30), VectorInput3(0.5, 0.5, 0.5), Deg * 90, 0, 0)
   Case 3: Scene(3).IDMatrix = MatrixWorld(VectorInput3(-20, -5, -10), VectorInput3(0.3, 0.3, 0.3), 0, 0, 0)
  End Select

  For I = LBound(Scene(J).Faces) To UBound(Scene(J).Faces)
   Select Case J
    Case 0: Scene(0).Faces(I).Col = ColorInput(50, 50, 50)
    Case 1: Scene(1).Faces(I).Col = ColorInput(150, 0, 0)
    Case 2: Scene(2).Faces(I).Col = ColorInput(0, 0, 100)
    Case 3: Scene(3).Faces(I).Col = ColorInput(0, 150, 0)
   End Select
   Scene(J).Faces(I).ShadVal = 0.5  'The scene can reflect 1/2 light (you can change: 0....1)
  Next I

 Next J

 'Do some stuffs (setup the camera and the spotlights):
 '=====================================================

 'Setup the camera:
 CamPos = VectorInput3(-5, -5.5, -30.5)      'Set Camera position
 ViewMatrix = MatrixIdentity                 'Set view matrix as identity
 ViewMode = False                            'Default value: Pitch/Yaw mode
 FogEnable = True                            'Enable Fogging
 LookAtObj = 1                               'Default value: At the Sphere
 XAng = 15

 'Setup the spotlights
 With Spots(0)                               'The first spot should be ambiant
  .Origin = VectorInput3(0.1, -10, 0.1)
  .Direction = VectorInput3(0.1, 0.1, 0.1)   'We avoid the zero
  .Falloff = 1: .Hotspot = 0.5
  .Color = ColorInput(150, 150, 150)         'Grey color
  .Ambiance = ColorInput(25, 25, 25)
  .DarkRange = FarPlane                      'Dark/Bright values (Atten enable only)
  .BrightRange = NearPlane
  .Diffusion = 1.5: .Specular = 0
  .AttenEnable = False: .Enabled = True
 End With

 With Spots(1)                               'The secend Spot
  .Origin = VectorInput3(80, -10, 80)
  .Direction = VectorInput3(0.1, 0.1, 0.1)   'We avoid the zero
  .Falloff = 1: .Hotspot = 0.5
  .Color = ColorInput(200, 200, 200)         'White color
  .Ambiance = ColorInput(0, 0, 0)
  .DarkRange = FarPlane                      'Dark/Bright values (Atten enable only)
  .BrightRange = NearPlane
  .Diffusion = 2: .Specular = 2              'Two Specular power
  .AttenEnable = False: .Enabled = True
 End With

 ReDim FacesDepth(0)                         'Init sort arrays
 ReDim FacesIndex(0)
 ReDim MeshsIndex(0)

End Sub
Sub Render()

 'Rasterization (2D):
 '===================
 ' Clip the faces and draw them with TCol (Shaded color).
 '
 ' In this app, Clipping is 2D, I want to add
 '  the 3D-Clipping in future versions.
 '
 ' Note that before rasterization, we split the clipped
 '  polygon into smalls triangles for rasterization (Triangulation)

 For I = LBound(FacesIndex) To UBound(FacesIndex)

  FInx = FacesIndex(I): MInx = MeshsIndex(I)

  X1 = 340 + Scene(MInx).TmpVerts(Scene(MInx).Faces(FInx).A).X
  Y1 = 260 + Scene(MInx).TmpVerts(Scene(MInx).Faces(FInx).A).Y
  X2 = 340 + Scene(MInx).TmpVerts(Scene(MInx).Faces(FInx).B).X
  Y2 = 260 + Scene(MInx).TmpVerts(Scene(MInx).Faces(FInx).B).Y
  X3 = 340 + Scene(MInx).TmpVerts(Scene(MInx).Faces(FInx).C).X
  Y3 = 260 + Scene(MInx).TmpVerts(Scene(MInx).Faces(FInx).C).Y

  FillColor = RGB(Scene(MInx).Faces(FInx).TCol.R, _
                  Scene(MInx).Faces(FInx).TCol.G, _
                  Scene(MInx).Faces(FInx).TCol.B)

  ClipTriangle 20, 20, 660, 500, X1, Y1, X2, Y2, X3, Y3, Pts(), NumPts, Stat

  Select Case Stat
   Case 1: 'Full triangle, so draw it normaly
           ReDim Pts(2)
           Pts(0).X = X1: Pts(0).Y = Y1
           Pts(1).X = X2: Pts(1).Y = Y2
           Pts(2).X = X3: Pts(2).Y = Y3
           Polygon hDC, Pts(0), 3

   Case 2: 'Polygon is clipped & triangulated
           ReDim Vs(1)
           For S = 0 To NumPts
            If S <> 0 Then ReDim Preserve Vs(UBound(Vs) + 1)
            Vs(UBound(Vs)).X = Pts(S).X
            Vs(UBound(Vs)).Y = Pts(S).Y
           Next S
           ReDim Preserve Vs(UBound(Vs) + 5)
           NumTris = Triangulate(CInt(UBound(Vs) - 5), Vs(), Ts())
           For S = 1 To NumTris
            Tri(0).X = Vs(Ts(S).A).X: Tri(0).Y = Vs(Ts(S).A).Y
            Tri(1).X = Vs(Ts(S).B).X: Tri(1).Y = Vs(Ts(S).B).Y
            Tri(2).X = Vs(Ts(S).C).X: Tri(2).Y = Vs(Ts(S).C).Y
            Polygon hDC, Tri(0), 3
           Next S
           ReDim Vs(1)
  End Select

 Next I

 ReDim FacesDepth(0)         'Clear sort arrays
 ReDim FacesIndex(0)
 ReDim MeshsIndex(0)

End Sub
Private Sub Form_Activate()

 '#############
 '# Main loop #
 '#############

 Do
  Cls

  GetKeys
  Process
  Render

  DoEvents
 Loop

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 'Generate the keyboard:
 '======================

 If KeyCode = vbKeyEscape Then KeyESC = True
 If KeyCode = vbKeySpace Then KeySPC = True
 If KeyCode = vbKeyLeft Then KeyLft = True
 If KeyCode = vbKeyRight Then KeyRgt = True
 If KeyCode = vbKeyUp Then KeyTop = True
 If KeyCode = vbKeyDown Then KeyBot = True
 If KeyCode = vbKeyHome Then KeyHom = True
 If KeyCode = vbKeyEnd Then KeyEnd = True
 If KeyCode = vbKeyF Then KeyF = True
 If KeyCode = vbKeyNumpad1 Then KeyPad1 = True
 If KeyCode = vbKeyNumpad2 Then KeyPad2 = True
 If KeyCode = vbKeyNumpad3 Then KeyPad3 = True
 If KeyCode = vbKeyNumpad4 Then KeyPad4 = True
 If KeyCode = vbKeyNumpad5 Then KeyPad5 = True

End Sub
Private Sub Form_Load()

 'Redim our window as (640x480), the 680x520
 'resolution is conceived for showing the clipping process (20, 20, 660, 500).

 Move 0, 0, (680 * 15), (520 * 15)
 ScaleMode = vbPixels

 LoadScene

End Sub
