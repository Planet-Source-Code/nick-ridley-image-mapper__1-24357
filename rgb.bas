Attribute VB_Name = "RGBCov"
Type COLORRGB
  red As Long
  green As Long
  blue As Long
End Type

Function RGB_get(ByVal CVal As Long, r As Long, B As Long, G As Long) As COLORRGB
G = Int(CVal / 65536)
B = Int((CVal - (65536 * G)) / 256)
r = CVal - (65536 * G + 256 * B)
End Function

Function Diff(Result As Integer, R1 As Long, R2 As Long, B1 As Long, B2 As Long, G1 As Long, G2 As Long, Tol As Long)
If R1 > R2 + Tol Or B1 > B2 + Tol Or G1 > G2 + Tol Or _
R1 < R2 - Tol Or B1 < B2 - Tol Or G1 < G2 - Tol Then
Result = 0 'False
Else
Result = 1 'True
End If
End Function
