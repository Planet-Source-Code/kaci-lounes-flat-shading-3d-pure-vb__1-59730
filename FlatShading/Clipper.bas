Attribute VB_Name = "Clipper"
Option Explicit
Function ClipDot(XX!, YY!, L!, T!, W!, H!) As Boolean

 'Check if a 2D point is inside a rectangle.

 If (XX > L) And (XX < W) Then
  If (YY > T) And (YY < H) Then ClipDot = True
 End If

End Function
Sub ClipTriangle(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!, X3!, Y3!, OutPts() As Point2D, NumPts As Byte, Stat As Byte)

 'Triangle clipping steps (by myself):
 '====================================
 '
 '1- Elimenate three cases (Completly inside, Region is small that triangle, Completly outside)
 '2- (Cases else): Clip line-by-line AB, BC, CA
 '3- Check if points region boundary is in this triangle, if yes,
 '    add this point to output array (there are four, but 3 as maximum):
 '    (RX1-RY1, RX1-RY2, RX2-RY1, RX2-RY2)
 '4- Remove the repeated points from output array.
 '5- Set points to clockwise order (Not included).
 '
 'Return: A polygon of NumPts points stored in OutPts() array.

 Dim A1%, A2%, A3%, R1%, R2%, R3%
 Dim OX1!, OY1!, OX2!, OY2!
 Dim I%, J%, K%

 'Get trivial cases:
 '---------------------------
 If Accept(RX1, RY1, RX2, RY2, X1, Y1, X2, Y2) = True Then A1 = -1 Else A1 = 0
 If Accept(RX1, RY1, RX2, RY2, X2, Y2, X3, Y3) = True Then A2 = -1 Else A2 = 0
 If Accept(RX1, RY1, RX2, RY2, X3, Y3, X1, Y1) = True Then A3 = -1 Else A3 = 0

 OX1 = 0: OY1 = 0: OX2 = 0: OY2 = 0
 ClipLine RX1, RY1, RX2, RY2, X1, Y1, X2, Y2, OX1, OY1, OX2, OY2
 If OX1 = 0 And OY1 = 0 And OX2 = 0 And OY2 = 0 Then R1 = -1

 OX1 = 0: OY1 = 0: OX2 = 0: OY2 = 0
 ClipLine RX1, RY1, RX2, RY2, X2, Y2, X3, Y3, OX1, OY1, OX2, OY2
 If OX1 = 0 And OY1 = 0 And OX2 = 0 And OY2 = 0 Then R2 = -1

 OX1 = 0: OY1 = 0: OX2 = 0: OY2 = 0
 ClipLine RX1, RY1, RX2, RY2, X3, Y3, X1, Y1, OX1, OY1, OX2, OY2
 If OX1 = 0 And OY1 = 0 And OX2 = 0 And OY2 = 0 Then R3 = -1
 '---------------------------

 'Completly inside:
 If A1 = 0 And A2 = 0 And A3 = 0 And _
    R1 = 0 And R2 = 0 And R3 = 0 Then
  ReDim OutPts(2)
  OutPts(0).X = X1: OutPts(0).Y = Y1
  OutPts(1).X = X2: OutPts(1).Y = Y2
  OutPts(2).X = X3: OutPts(2).Y = Y3
  NumPts = 2: Stat = 1
  Exit Sub
 End If

 'Region is in the big triangle:
 If IsPtrInTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, RX1, RY1) = True And _
    IsPtrInTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, RX2, RY1) = True And _
    IsPtrInTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, RX2, RY2) = True And _
    IsPtrInTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, RX1, RY2) = True Then
  ReDim OutPts(3)
  OutPts(0).X = RX1: OutPts(0).Y = RY1
  OutPts(1).X = RX2: OutPts(1).Y = RY1
  OutPts(2).X = RX2: OutPts(2).Y = RY2
  OutPts(3).X = RX1: OutPts(3).Y = RY2
  NumPts = 3
  Stat = 2
  Exit Sub
 End If

 'Completly outside:
 If A1 = -1 And A2 = -1 And A3 = -1 And _
    R1 = -1 And R2 = -1 And R3 = -1 Then NumPts = 0: Stat = 0: Exit Sub

 '****************************************************************************

 Stat = 2
 ReDim OutPts(0)

 'Edge1: AB -----------------------------------------------------------
 If A1 = -1 Then
  If R1 = 0 Then
   ReDim OutPts(1)
   ClipLine RX1, RY1, RX2, RY2, X1, Y1, X2, Y2, OX1, OY1, OX2, OY2
   OutPts(0).X = OX1: OutPts(0).Y = OY1
   OutPts(1).X = OX2: OutPts(1).Y = OY2
  End If
 Else
  ReDim OutPts(1)
  OutPts(0).X = X1: OutPts(0).Y = Y1
  OutPts(1).X = X2: OutPts(1).Y = Y2
 End If

 'Edge2: BC -----------------------------------------------------------
 If A2 = -1 Then
  If R2 = 0 Then
   ClipLine RX1, RY1, RX2, RY2, X2, Y2, X3, Y3, OX1, OY1, OX2, OY2
   If UBound(OutPts) = 0 Then
    ReDim OutPts(1)
    OutPts(0).X = OX1: OutPts(0).Y = OY1
    OutPts(1).X = OX2: OutPts(1).Y = OY2
   Else
    ReDim Preserve OutPts(UBound(OutPts) + 2)
    OutPts(UBound(OutPts) - 1).X = OX1
    OutPts(UBound(OutPts) - 1).Y = OY1
    OutPts(UBound(OutPts)).X = OX2
    OutPts(UBound(OutPts)).Y = OY2
   End If
  End If
 Else
  If UBound(OutPts) = 0 Then
   ReDim OutPts(1)
   OutPts(0).X = X2: OutPts(0).Y = Y2
   OutPts(1).X = X3: OutPts(1).Y = Y3
  Else
   ReDim Preserve OutPts(UBound(OutPts) + 2)
   OutPts(UBound(OutPts) - 1).X = X2
   OutPts(UBound(OutPts) - 1).Y = Y2
   OutPts(UBound(OutPts)).X = X3
   OutPts(UBound(OutPts)).Y = Y3
  End If
 End If

 'Edge3: CA -----------------------------------------------------------
 If A3 = -1 Then
  If R3 = 0 Then
   ClipLine RX1, RY1, RX2, RY2, X3, Y3, X1, Y1, OX1, OY1, OX2, OY2
   If UBound(OutPts) = 0 Then
    ReDim OutPts(1)
    OutPts(0).X = OX1: OutPts(0).Y = OY1
    OutPts(1).X = OX2: OutPts(1).Y = OY2
   Else
    ReDim Preserve OutPts(UBound(OutPts) + 2)
    OutPts(UBound(OutPts) - 1).X = OX1
    OutPts(UBound(OutPts) - 1).Y = OY1
    OutPts(UBound(OutPts)).X = OX2
    OutPts(UBound(OutPts)).Y = OY2
   End If
  End If
 Else
  If UBound(OutPts) = 0 Then
   ReDim OutPts(1)
   OutPts(0).X = X3: OutPts(0).Y = Y3
   OutPts(1).X = X1: OutPts(1).Y = Y1
  Else
   ReDim Preserve OutPts(UBound(OutPts) + 2)
   OutPts(UBound(OutPts) - 1).X = X3
   OutPts(UBound(OutPts) - 1).Y = Y3
   OutPts(UBound(OutPts)).X = X1
   OutPts(UBound(OutPts)).Y = Y1
  End If
 End If

'-----------------------------------------------------------------------

 If IsPtrInTriangle(X1, Y1, X2, Y2, X3, Y3, RX1, RY1) = True Then
  ReDim Preserve OutPts(UBound(OutPts) + 1)
  OutPts(UBound(OutPts)).X = RX1
  OutPts(UBound(OutPts)).Y = RY1
 End If

 If IsPtrInTriangle(X1, Y1, X2, Y2, X3, Y3, RX2, RY1) = True Then
  ReDim Preserve OutPts(UBound(OutPts) + 1)
  OutPts(UBound(OutPts)).X = RX2
  OutPts(UBound(OutPts)).Y = RY1
 End If

 If IsPtrInTriangle(X1, Y1, X2, Y2, X3, Y3, RX2, RY2) = True Then
  ReDim Preserve OutPts(UBound(OutPts) + 1)
  OutPts(UBound(OutPts)).X = RX2
  OutPts(UBound(OutPts)).Y = RY2
 End If

 If IsPtrInTriangle(X1, Y1, X2, Y2, X3, Y3, RX1, RY2) = True Then
  ReDim Preserve OutPts(UBound(OutPts) + 1)
  OutPts(UBound(OutPts)).X = RX1
  OutPts(UBound(OutPts)).Y = RY2
 End If

 'Remove repeated points:
ReCheck:
 For I = LBound(OutPts) To UBound(OutPts)
  For J = I To UBound(OutPts)
   If (OutPts(I).X = OutPts(J).X) And (OutPts(I).Y = OutPts(J).Y) And (J <> I) Then
    For K = I To UBound(OutPts) - 1
     OutPts(K).X = OutPts(K + 1).X
     OutPts(K).Y = OutPts(K + 1).Y
    Next K
    ReDim Preserve OutPts(UBound(OutPts) - 1)
    GoTo ReCheck
   End If
  Next J
 Next I

 NumPts = UBound(OutPts)

End Sub
Function IsPtrInTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, XX!, YY!) As Boolean

 'Check if a 2D point is inside a triangle (Barycentric).

 Dim BC!, CA!, AB!, AP!, BP!, CP!, ABC!

 AB = (X1 * Y2) - (Y1 * X2)
 BC = (X2 * Y3) - (Y2 * X3)
 CA = (X3 * Y1) - (Y3 * X1)

 AP = (X1 * YY) - (Y1 * XX)
 BP = (X2 * YY) - (Y2 * XX)
 CP = (X3 * YY) - (Y3 * XX)

 ABC = (BC + CA + AB)

 If ABC < 0 Then
  ABC = -1
 ElseIf ABC = 0 Then
  ABC = 0
 ElseIf ABC > 0 Then
  ABC = 1
 End If

 If (ABC * (BC - BP + CP) > 0) And _
    (ABC * (CA - CP + AP) > 0) And _
    (ABC * (AB - AP + BP) > 0) Then IsPtrInTriangle = True

End Function
Sub ClipLine(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!, OutX1!, OutY1!, OutX2!, OutY2!)

 'Liang Barsky Line Clipping Algorithm (1984):
 ' (Parametric clipping but special case for rectangular clipping regions)
 '
 'Note that for fast checking the trivial cases, I use Cohen-Sutherland (codes),
 ' But for clipping the line, I use directly Liang-Barsky algorithm (without code)
 '  Look the first line in this routine:
 '
 '   RX1, RY1, RX2, RY2 : The region coordinates (rectangle region)
 '   X1, Y1, X2, Y2     : The coordinates of the line to be clipped (Input)
 '   OX1, OY1, OX2, OY2 : The coordinates of the clipped line (Output)
 '
 'Also note that routine is *Very* optimized!
 ' Realy, it is two modules that i have it transformed in only one procedure!!!!!

 Dim PX1!, PY1!, PX2!, PY2!
 Dim TX1!, TY1!, TX2!, TY2!
 Dim U1!, U2!, Dx!, Dy!, Temp!
 Dim P!, Q!, UU1!, UU2!, R!, CT As Byte

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 PX1 = X1: PY1 = Y1: PX2 = X2: PY2 = Y2
 U1 = 0: U2 = 1: Dx = (PX2 - PX1)

 P = (-1 * Dx): Q = (PX1 - RX1): UU1 = U1: UU2 = U1: CT = 1
 If P < 0 Then
  R = Q / P: If R > U2 Then CT = 0 Else If R > U1 Then U1 = R
 Else
  If P > 0 Then
   R = Q / P: If R < U1 Then CT = 0 Else If R < U2 Then U2 = R
  ElseIf Q < 0 Then
   CT = 0
  End If
 End If

 If CT = 1 Then
  P = Dx: Q = (RX2 - PX1): UU1 = U1: UU2 = U2: CT = 1
  If P < 0 Then
   R = Q / P: If R > U2 Then CT = 0 Else If R > U1 Then U1 = R
  Else
   If P > 0 Then
    R = Q / P: If R < U1 Then CT = 0 Else If R < U2 Then U2 = R
   ElseIf Q < 0 Then
    CT = 0
   End If
  End If
  If CT = 1 Then
   Dy = (PY2 - PY1): P = (-1 * Dy): Q = (PY1 - RY1): UU1 = U1: UU2 = U2: CT = 1
   If P < 0 Then
    R = Q / P: If R > U2 Then CT = 0 Else If R > U1 Then U1 = R
   Else
    If P > 0 Then
     R = Q / P: If R < U1 Then CT = 0 Else If R < U2 Then U2 = R
    ElseIf Q < 0 Then
     CT = 0
    End If
   End If
   If CT = 1 Then
    P = Dy: Q = (RY2 - PY1): UU1 = U1: UU2 = U2: CT = 1
    If P < 0 Then
     R = Q / P: If R > U2 Then CT = 0 Else If R > U1 Then U1 = R
    Else
     If P > 0 Then
      R = Q / P: If R < U1 Then CT = 0 Else If R < U2 Then U2 = R
     ElseIf Q < 0 Then
      CT = 0
     End If
    End If
    If CT = 1 Then
     If U2 < 1 Then PX2 = PX1 + (U2 * Dx): PY2 = PY1 + (U2 * Dy)
     If U1 > 0 Then PX1 = PX1 + (U1 * Dx): PY1 = PY1 + (U1 * Dy)
     OutX1 = PX1: OutY1 = PY1: OutX2 = PX2: OutY2 = PY2
    End If
   End If
  End If
 End If

End Sub
Function Accept(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!) As Boolean

 'Cohen-Sutherland Trivial Accept (with codes):

 Dim Code1(3) As Boolean
 Dim Code2(3) As Boolean
 Dim Temp As Single

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 If X1 < RX1 Then Code1(0) = True Else Code1(0) = False
 If X1 > RX2 Then Code1(1) = True Else Code1(1) = False
 If Y1 < RY1 Then Code1(2) = True Else Code1(2) = False
 If Y1 > RY2 Then Code1(3) = True Else Code1(3) = False

 If X2 < RX1 Then Code2(0) = True Else Code2(0) = False
 If X2 > RX2 Then Code2(1) = True Else Code2(1) = False
 If Y2 < RY1 Then Code2(2) = True Else Code2(2) = False
 If Y2 > RY2 Then Code2(3) = True Else Code2(3) = False

 If (Code1(0) Or Code2(0)) Then Accept = True
 If (Code1(1) Or Code2(1)) Then Accept = True
 If (Code1(2) Or Code2(2)) Then Accept = True
 If (Code1(3) Or Code2(3)) Then Accept = True

End Function
Function Reject(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!) As Boolean

 'Cohen-Sutherland Trivial Reject (with codes):

 Dim Code1(3) As Boolean
 Dim Code2(3) As Boolean
 Dim Temp As Single

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 If X1 < RX1 Then Code1(0) = True Else Code1(0) = False
 If X1 > RX2 Then Code1(1) = True Else Code1(1) = False
 If Y1 < RY1 Then Code1(2) = True Else Code1(2) = False
 If Y1 > RY2 Then Code1(3) = True Else Code1(3) = False

 If X2 < RX1 Then Code2(0) = True Else Code2(0) = False
 If X2 > RX2 Then Code2(1) = True Else Code2(1) = False
 If Y2 < RY1 Then Code2(2) = True Else Code2(2) = False
 If Y2 > RY2 Then Code2(3) = True Else Code2(3) = False

 If (Code1(0) And Code2(0)) Then Reject = True
 If (Code1(1) And Code2(1)) Then Reject = True
 If (Code1(2) And Code2(2)) Then Reject = True
 If (Code1(3) And Code2(3)) Then Reject = True

End Function
