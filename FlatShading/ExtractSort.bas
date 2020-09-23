Attribute VB_Name = "ExtractSort"
Option Explicit
Sub ExtractSort3D(FacesDepth() As Single, FacesIndex() As Long, MeshsIndex() As Byte)

 'Few code, more speed !!!
 '
 'Extraction Sort: Very very simple algorithm,
 ' in 10 lines of code only !, and 2 times faster than QuickSort !
 '
 'Note: Is this program, the routine sort an array of an
 '      averaged depths of faces (Zs/3), in the same time, change
 '      the indexs of faces in FacesIndex array (also in MeshIndex array),
 '      Then, draw faces from lower boundary to upper boundary of this last.
 '      contrarily to sort the faces them even, that requires
 '      more CPU time, I prefer to sort an array of 'Singles' (4 Bytes)
 '      that sort an array of 'Face' data type (46 Bytes).

 Dim I&, J&, K&, TmpDpth!, TmpFace%, TmpMesh As Byte

 For I = LBound(FacesDepth) To UBound(FacesDepth)

  K = I

  For J = I To UBound(FacesDepth)
   If FacesDepth(K) < FacesDepth(J) Then K = J
  Next J

  TmpDpth = FacesDepth(K): FacesDepth(K) = FacesDepth(I): FacesDepth(I) = TmpDpth
  TmpFace = FacesIndex(K): FacesIndex(K) = FacesIndex(I): FacesIndex(I) = TmpFace
  TmpMesh = MeshsIndex(K): MeshsIndex(K) = MeshsIndex(I): MeshsIndex(I) = TmpMesh

 Next I

End Sub
