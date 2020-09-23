Attribute VB_Name = "Maths3D"
Option Explicit

Global Const Pi! = 3.141593
Global Const Deg! = (Pi / 180)

Public Declare Function Polygon Lib "Gdi32.dll" (ByVal hDC As Long, lpPoint As Point2D, ByVal nCount As Long) As Long

Public Type Matrix    '(128 Bytes)
 M11 As Single: M12 As Single: M13 As Single: M14 As Single
 M21 As Single: M22 As Single: M23 As Single: M24 As Single
 M31 As Single: M32 As Single: M33 As Single: M34 As Single
 M41 As Single: M42 As Single: M43 As Single: M44 As Single
End Type

Public Type Vector3D  '(12 Bytes)
 X As Single
 Y As Single
 Z As Single
End Type

Public Type Point2D   '(8 Bytes)
 X As Long
 Y As Long
End Type

Public Type ColRGB    '(6 Bytes)
 R As Integer
 G As Integer
 B As Integer
End Type

'#########################################

Public Type Face      '(46 Bytes)
 A As Integer
 B As Integer
 C As Integer
 Normal As Vector3D
 Center As Vector3D
 ShadVal As Single
 Col As ColRGB
 TCol As ColRGB
End Type

Public Type Mesh
 Vertices() As Vector3D
 TmpVerts() As Vector3D
 Faces() As Face
 IDMatrix As Matrix
End Type

Public Type SpotLight3D  '(64 Bytes)
 Origin As Vector3D       '------- Parametres ------
 Direction As Vector3D
 TPos As Vector3D
 TDir As Vector3D
 Color As ColRGB
 Falloff As Single
 Hotspot As Single
 BrightRange As Single
 DarkRange As Single
 Enabled As Boolean
 Ambiance As ColRGB       '------- Materials -------
 Diffusion As Single
 Specular As Single
 AttenEnable As Boolean
End Type
Function ColorInput(R%, G%, B%) As ColRGB

 ColorInput.R = R
 ColorInput.G = G
 ColorInput.B = B

End Function
Function VectorGetCenter(VecA As Vector3D, VecB As Vector3D, VecC As Vector3D) As Vector3D

 VectorGetCenter.X = (VecA.X + VecB.X + VecC.X) * 0.3333333
 VectorGetCenter.Y = (VecA.Y + VecB.Y + VecC.Y) * 0.3333333
 VectorGetCenter.Z = (VecA.Z + VecB.Z + VecC.Z) * 0.3333333

End Function
Function VectorDistance3(VecA As Vector3D, VecB As Vector3D) As Single

 VectorDistance3 = VectorLength3(VectorSubtract3(VecA, VecB))

End Function
Function VectorReflect(VecA As Vector3D, VecB As Vector3D) As Vector3D

 If VectorAngle(VecA, VecB) < 0 Then
  VectorReflect = VectorAdd3(VecA, VectorScale3(VectorScale3(VectorNormalize(VecB), VectorDotProduct3(VecA, VectorNormalize(VecB))), -2))
 End If

End Function
Function ColorAdd(A As ColRGB, B As ColRGB) As ColRGB

 ColorAdd.R = (A.R + B.R)
 ColorAdd.G = (A.G + B.G)
 ColorAdd.B = (A.B + B.B)

End Function
Function ColorInterpolate(A As ColRGB, B As ColRGB, Alpha As Single) As ColRGB

 ColorInterpolate.R = ((B.R - A.R) * Alpha) + A.R
 ColorInterpolate.G = ((B.G - A.G) * Alpha) + A.G
 ColorInterpolate.B = ((B.B - A.B) * Alpha) + A.B

End Function
Function ColorScale(A As ColRGB, B As Single) As ColRGB

 ColorScale.R = (A.R * B)
 ColorScale.G = (A.G * B)
 ColorScale.B = (A.B * B)

End Function
Function VectorAngle(VecA As Vector3D, VecB As Vector3D) As Single

 If VectorCompare3(VecA, VectorNull3) = False And VectorCompare3(VecB, VectorNull3) = False Then
  VectorAngle = VectorDotProduct3(VectorNormalize(VecA), VectorNormalize(VecB))
 End If

End Function
Function VectorCrossProduct(VecA As Vector3D, VecB As Vector3D) As Vector3D

 VectorCrossProduct.X = (VecA.Y * VecB.Z) - (VecA.Z * VecB.Y)
 VectorCrossProduct.Y = (VecA.Z * VecB.X) - (VecA.X * VecB.Z)
 VectorCrossProduct.Z = (VecA.X * VecB.Y) - (VecA.Y * VecB.X)

End Function
Function VectorDotProduct3(VecA As Vector3D, VecB As Vector3D) As Single

 VectorDotProduct3 = (VecA.X * VecB.X) + (VecA.Y * VecB.Y) + (VecA.Z * VecB.Z)

End Function
Function VectorGetNormal(VecA As Vector3D, VecB As Vector3D, VecC As Vector3D) As Vector3D

 VectorGetNormal = VectorCrossProduct(VectorSubtract3(VecA, VecB), VectorSubtract3(VecC, VecB))

End Function
Function VectorInput3(X!, Y!, Z!) As Vector3D

 VectorInput3.X = X
 VectorInput3.Y = Y
 VectorInput3.Z = Z

End Function
Function VectorLength3(Vec As Vector3D) As Single

 VectorLength3 = Sqr((Vec.X * Vec.X) + (Vec.Y * Vec.Y) + (Vec.Z * Vec.Z))

End Function
Function VectorNormalize(Vec As Vector3D) As Vector3D

 If VectorCompare3(Vec, VectorNull3) = False Then

  Dim L As Single: L = (1 / VectorLength3(Vec))

  VectorNormalize.X = (Vec.X * L)
  VectorNormalize.Y = (Vec.Y * L)
  VectorNormalize.Z = (Vec.Z * L)
 
 End If

End Function
Function VectorCompare3(VecA As Vector3D, VecB As Vector3D) As Boolean

 If (VecA.X = VecB.X) And (VecA.Y = VecB.Y) And (VecA.Z = VecB.Z) Then VectorCompare3 = True

End Function
Function VectorNull3() As Vector3D

End Function
Function VectorRotate(Vec As Vector3D, Axis As Byte, Angle As Single) As Vector3D

 'Basic rotations (without matrices)

 Select Case Axis
  Case 0:
   VectorRotate.X = Vec.X
   VectorRotate.Y = (Cos(Angle) * Vec.Y) - (Sin(Angle) * Vec.Z)
   VectorRotate.Z = (Sin(Angle) * Vec.Y) + (Cos(Angle) * Vec.Z)
  Case 1:
   VectorRotate.X = (Cos(Angle) * Vec.X) + (Sin(Angle) * Vec.Z)
   VectorRotate.Y = Vec.Y
   VectorRotate.Z = -(Sin(Angle) * Vec.X) + (Cos(Angle) * Vec.Z)
  Case 2:
   VectorRotate.X = (Cos(Angle) * Vec.X) - (Sin(Angle) * Vec.Y)
   VectorRotate.Y = (Sin(Angle) * Vec.X) + (Cos(Angle) * Vec.Y)
   VectorRotate.Z = Vec.Z
 End Select

End Function
Function VectorAdd3(VecA As Vector3D, VecB As Vector3D) As Vector3D

 VectorAdd3.X = (VecA.X + VecB.X)
 VectorAdd3.Y = (VecA.Y + VecB.Y)
 VectorAdd3.Z = (VecA.Z + VecB.Z)

End Function
Function VectorScale3(Vec As Vector3D, S As Single) As Vector3D

 VectorScale3.X = (Vec.X * S)
 VectorScale3.Y = (Vec.Y * S)
 VectorScale3.Z = (Vec.Z * S)

End Function
Function VectorSubtract3(VecA As Vector3D, VecB As Vector3D) As Vector3D

 VectorSubtract3.X = (VecA.X - VecB.X)
 VectorSubtract3.Y = (VecA.Y - VecB.Y)
 VectorSubtract3.Z = (VecA.Z - VecB.Z)

End Function
Function MatrixIdentity() As Matrix

 With MatrixIdentity
  .M11 = 1: .M12 = 0: .M13 = 0: .M14 = 0
  .M21 = 0: .M22 = 1: .M23 = 0: .M24 = 0
  .M31 = 0: .M32 = 0: .M33 = 1: .M34 = 0
  .M41 = 0: .M42 = 0: .M43 = 0: .M44 = 1
 End With

End Function
Function MatrixMultiply(MatA As Matrix, MatB As Matrix) As Matrix

 With MatrixMultiply
  .M11 = (MatA.M11 * MatB.M11) + (MatA.M21 * MatB.M12) + (MatA.M31 * MatB.M13) + (MatA.M41 * MatB.M14)
  .M12 = (MatA.M12 * MatB.M11) + (MatA.M22 * MatB.M12) + (MatA.M32 * MatB.M13) + (MatA.M42 * MatB.M14)
  .M13 = (MatA.M13 * MatB.M11) + (MatA.M23 * MatB.M12) + (MatA.M33 * MatB.M13) + (MatA.M43 * MatB.M14)
  .M14 = (MatA.M14 * MatB.M11) + (MatA.M24 * MatB.M12) + (MatA.M34 * MatB.M13) + (MatA.M44 * MatB.M14)
  .M21 = (MatA.M11 * MatB.M21) + (MatA.M21 * MatB.M22) + (MatA.M31 * MatB.M23) + (MatA.M41 * MatB.M24)
  .M22 = (MatA.M12 * MatB.M21) + (MatA.M22 * MatB.M22) + (MatA.M32 * MatB.M23) + (MatA.M42 * MatB.M24)
  .M23 = (MatA.M13 * MatB.M21) + (MatA.M23 * MatB.M22) + (MatA.M33 * MatB.M23) + (MatA.M43 * MatB.M24)
  .M24 = (MatA.M14 * MatB.M21) + (MatA.M24 * MatB.M22) + (MatA.M34 * MatB.M23) + (MatA.M44 * MatB.M24)
  .M31 = (MatA.M11 * MatB.M31) + (MatA.M21 * MatB.M32) + (MatA.M31 * MatB.M33) + (MatA.M41 * MatB.M34)
  .M32 = (MatA.M12 * MatB.M31) + (MatA.M22 * MatB.M32) + (MatA.M32 * MatB.M33) + (MatA.M42 * MatB.M34)
  .M33 = (MatA.M13 * MatB.M31) + (MatA.M23 * MatB.M32) + (MatA.M33 * MatB.M33) + (MatA.M43 * MatB.M34)
  .M34 = (MatA.M14 * MatB.M31) + (MatA.M24 * MatB.M32) + (MatA.M34 * MatB.M33) + (MatA.M44 * MatB.M34)
  .M41 = (MatA.M11 * MatB.M41) + (MatA.M21 * MatB.M42) + (MatA.M31 * MatB.M43) + (MatA.M41 * MatB.M44)
  .M42 = (MatA.M12 * MatB.M41) + (MatA.M22 * MatB.M42) + (MatA.M32 * MatB.M43) + (MatA.M42 * MatB.M44)
  .M43 = (MatA.M13 * MatB.M41) + (MatA.M23 * MatB.M42) + (MatA.M33 * MatB.M43) + (MatA.M43 * MatB.M44)
  .M44 = (MatA.M14 * MatB.M41) + (MatA.M24 * MatB.M42) + (MatA.M34 * MatB.M43) + (MatA.M44 * MatB.M44)
 End With

End Function
Function MatrixMultiplyVector3(Vec As Vector3D, Mat As Matrix) As Vector3D

 MatrixMultiplyVector3.X = (Mat.M11 * Vec.X) + (Mat.M12 * Vec.Y) + (Mat.M13 * Vec.Z) + (Mat.M14)
 MatrixMultiplyVector3.Y = (Mat.M21 * Vec.X) + (Mat.M22 * Vec.Y) + (Mat.M23 * Vec.Z) + (Mat.M24)
 MatrixMultiplyVector3.Z = (Mat.M31 * Vec.X) + (Mat.M32 * Vec.Y) + (Mat.M33 * Vec.Z) + (Mat.M34)

End Function
Function MatrixRotate(Axis As Byte, Angle As Single) As Matrix

 Select Case Axis
  Case 0:
   With MatrixRotate
    .M11 = 1
    .M22 = Cos(Angle)
    .M23 = -Sin(Angle)
    .M32 = -.M23
    .M33 = .M22
    .M44 = 1
   End With
  Case 1:
   With MatrixRotate
    .M11 = Cos(Angle)
    .M13 = Sin(Angle)
    .M22 = 1
    .M31 = -.M13
    .M33 = .M11
    .M44 = 1
   End With
  Case 2:
   With MatrixRotate
    .M11 = Cos(Angle)
    .M12 = -Sin(Angle)
    .M21 = -.M12
    .M22 = .M11
    .M33 = 1
    .M44 = 1
   End With
 End Select

End Function
Function MatrixScale3(Factor As Vector3D) As Matrix

 MatrixScale3.M11 = Factor.X
 MatrixScale3.M22 = Factor.Y
 MatrixScale3.M33 = Factor.Z
 MatrixScale3.M44 = 1

End Function
Function MatrixTranslate(Distance As Vector3D) As Matrix

 MatrixTranslate.M11 = 1
 MatrixTranslate.M14 = Distance.X
 MatrixTranslate.M22 = 1
 MatrixTranslate.M24 = Distance.Y
 MatrixTranslate.M33 = 1
 MatrixTranslate.M34 = Distance.Z
 MatrixTranslate.M44 = 1

End Function
Function MatrixView(VecFrom As Vector3D, VecLookAt As Vector3D, VecUp As Vector3D) As Matrix

 Dim MatRotat As Matrix, MatTrans As Matrix
 Dim U As Vector3D, V As Vector3D, N As Vector3D

 MatTrans = MatrixTranslate(VectorInput3(-VecFrom.X, -VecFrom.Y, -VecFrom.Z))

 N = VectorNormalize(VectorSubtract3(VecLookAt, VecFrom))
 U = VectorNormalize(VectorCrossProduct(VecUp, N))
 V = VectorCrossProduct(N, U)

 With MatRotat
  .M11 = U.X: .M12 = U.Y: .M13 = U.Z
  .M21 = V.X: .M22 = V.Y: .M23 = V.Z
  .M31 = N.X: .M32 = N.Y: .M33 = N.Z
  .M41 = 0: .M42 = 0: .M43 = 0: .M44 = 1
 End With

 MatrixView = MatrixIdentity
 MatrixView = MatrixMultiply(MatrixView, MatTrans)
 MatrixView = MatrixMultiply(MatrixView, MatRotat)

End Function
Function MatrixWorld(VecTranslate As Vector3D, VecScale As Vector3D, XPitch!, YYaw!, ZRoll!) As Matrix

 Dim MatTrans As Matrix, MatRotat As Matrix, MatScale As Matrix

 MatTrans = MatrixTranslate(VecTranslate)
 MatScale = MatrixScale3(VecScale)

 MatRotat = MatrixRotate(0, XPitch)
 MatRotat = MatrixMultiply(MatRotat, MatrixRotate(1, YYaw))
 MatRotat = MatrixMultiply(MatRotat, MatrixRotate(2, ZRoll))

 MatrixWorld = MatrixMultiply(MatrixIdentity, MatScale)
 MatrixWorld = MatrixMultiply(MatrixWorld, MatRotat)
 MatrixWorld = MatrixMultiply(MatrixWorld, MatTrans)

End Function
