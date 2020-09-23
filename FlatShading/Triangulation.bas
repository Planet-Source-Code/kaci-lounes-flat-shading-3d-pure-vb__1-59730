Attribute VB_Name = "Triangulation"
Option Explicit

Type Triangle  '(6 Bytes)
 A As Integer
 B As Integer
 C As Integer
End Type
Function InCircumCircle(XX&, YY&, X1&, Y1&, X2&, Y2&, X3&, Y3&, CX!, CY!, R!) As Boolean

 'Return: True if the point (XX, YY) lies inside the CircumCircle
 '         made up by points (X1, Y1) (X2, Y2) (X3, Y3).
 '        The CircumCircle centre is returned in (CX, CY) and the radius as R.

 Const Eps! = 0.000001

 Dim RSqr!, DRSqr!, Dx!, Dy!
 Dim M1!, M2!, MX1!, MX2!, MY1!, MY2!

 If (Abs(Y1 - Y2) < Eps) And (Abs(Y2 - Y3) < Eps) Then Exit Function

 If Abs(Y2 - Y1) < Eps Then
  M2 = -(X3 - X2) / (Y3 - Y2)
  MX2 = (X2 + X3) * 0.5: MY2 = (Y2 + Y3) * 0.5
  CX = (X2 + X1) * 0.5: CY = (M2 * (CX - MX2)) + MY2
 ElseIf Abs(Y3 - Y2) < Eps Then
  M1 = -(X2 - X1) / (Y2 - Y1)
  MX1 = (X1 + X2) * 0.5: MY1 = (Y1 + Y2) * 0.5
  CX = (X3 + X2) * 0.5: CY = (M1 * (CX - MX1)) + MY1
 Else
  M1 = -(X2 - X1) / (Y2 - Y1): M2 = -(X3 - X2) / (Y3 - Y2)
  MX1 = (X1 + X2) * 0.5: MX2 = (X2 + X3) * 0.5
  MY1 = (Y1 + Y2) * 0.5: MY2 = (Y2 + Y3) * 0.5
  CX = (((M1 * MX1) - (M2 * MX2)) + (MY2 - MY1)) / (M1 - M2)
  CY = (M1 * (CX - MX1)) + MY1
 End If

 Dx = (X2 - CX): Dy = (Y2 - CY)
 RSqr = (Dx * Dx) + (Dy * Dy): R = Sqr(RSqr)
 Dx = (XX - CX): Dy = (YY - CY)
 DRSqr = (Dx * Dx) + (Dy * Dy)

 If DRSqr <= RSqr Then InCircumCircle = True

End Function
Function Triangulate(NVert%, Verts() As Point2D, Tris() As Triangle) As Integer

 'Optimized Delaunay triangulation procedure:
 '===========================================
 '
 'Action: Takes as input NVert% vertices in array Verts(),
 '         A list of triangulated faces is stored in the Tris() array,
 '          These triangles are arranged in clockwise order.
 'Return: Number of triangles.

 Dim Edges() As Integer
 Dim Complete() As Boolean
 Dim XMin&, XMax&, YMin&, YMax&, XMid&, YMid&
 Dim I%, J%, K%, NTri%, NEdge&, Inc As Boolean
 Dim Dx!, Dy!, DMax!, XC!, YC!, R!

 ReDim Complete(UBound(Tris))
 ReDim Edges(2, (UBound(Tris) * 3))

 XMin = Verts(1).X: YMin = Verts(1).Y
 XMax = XMin: YMax = YMin

 For I = 2 To NVert
  If Verts(I).X < XMin Then XMin = Verts(I).X
  If Verts(I).X > XMax Then XMax = Verts(I).X
  If Verts(I).Y < YMin Then YMin = Verts(I).Y
  If Verts(I).Y > YMax Then YMax = Verts(I).Y
 Next I

 Dx = (XMax - XMin): Dy = (YMax - YMin)
 If (Dx > Dy) Then DMax = Dx Else DMax = Dy

 XMid = (XMax + XMin) * 0.5
 YMid = (YMax + YMin) * 0.5

 Verts(NVert + 1).X = XMid - (2 * DMax)
 Verts(NVert + 1).Y = (YMid - DMax)
 Verts(NVert + 2).X = XMid
 Verts(NVert + 2).Y = YMid + (2 * DMax)
 Verts(NVert + 3).X = XMid + (2 * DMax)
 Verts(NVert + 3).Y = (YMid - DMax)

 Tris(1).A = (NVert + 1): Tris(1).B = (NVert + 2): Tris(1).C = (NVert + 3)
 Complete(1) = False: NTri = 1

 For I = 1 To NVert
  NEdge = 0: J = 0
  Do
   J = (J + 1)
   If (Complete(J) = False) Then
    Inc = InCircumCircle(Verts(I).X, Verts(I).Y, Verts(Tris(J).A).X, Verts(Tris(J).A).Y, Verts(Tris(J).B).X, Verts(Tris(J).B).Y, Verts(Tris(J).C).X, Verts(Tris(J).C).Y, XC, YC, R)
    If (Inc = True) Then
     Edges(1, NEdge + 1) = Tris(J).A: Edges(2, NEdge + 1) = Tris(J).B
     Edges(1, NEdge + 2) = Tris(J).B: Edges(2, NEdge + 2) = Tris(J).C
     Edges(1, NEdge + 3) = Tris(J).C: Edges(2, NEdge + 3) = Tris(J).A
     Tris(J).A = Tris(NTri).A: Tris(J).B = Tris(NTri).B: Tris(J).C = Tris(NTri).C
     Complete(J) = Complete(NTri)
     NEdge = (NEdge + 3): J = (J - 1): NTri = (NTri - 1)
    End If
   End If
  Loop While (J < NTri)

  For J = 1 To (NEdge - 1)
   If (Edges(1, J) <> 0) And (Edges(2, J) <> 0) Then
    For K = (J + 1) To NEdge
     If (Edges(1, K) <> 0) And (Edges(2, K) <> 0) Then
      If (Edges(1, J) = Edges(2, K)) Then
       If (Edges(2, J) = Edges(1, K)) Then
        Edges(1, J) = 0: Edges(2, J) = 0
        Edges(1, K) = 0: Edges(2, K) = 0
       End If
      End If
     End If
    Next K
   End If
  Next J

  For J = 1 To NEdge
   If (Edges(1, J) <> 0) And (Edges(2, J) <> 0) Then
    NTri = (NTri + 1)
    Tris(NTri).A = Edges(1, J)
    Tris(NTri).B = Edges(2, J)
    Tris(NTri).C = I
    Complete(NTri) = False
   End If
  Next J

 Next I

 I = 0
 Do
  I = (I + 1)
  If (Tris(I).A > NVert) Or (Tris(I).B > NVert) Or (Tris(I).C > NVert) Then
   Tris(I).A = Tris(NTri).A: Tris(I).B = Tris(NTri).B: Tris(I).C = Tris(NTri).C
   I = (I - 1): NTri = (NTri - 1)
  End If
 Loop While (I < NTri)

 Triangulate = NTri

End Function
