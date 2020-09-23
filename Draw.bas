Attribute VB_Name = "Module1"
Option Explicit
Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Global LastX As Long
Global LastY As Long
Global DemoMode As Boolean
Global PunktCounter As Long
Global DoExplosion As Boolean

Type Datentyp_Punkt
  Active As Boolean
  X As Long
  Y As Long
  Color As Long
  StepX As Long
  StepY As Long
  LastX As Long
  LastY As Long
  FallMax As Long
End Type

Type Datentyp_Explosion
  DoIt As Boolean
  Great As Boolean
  Count As Long
End Type

Global Explosion As Datentyp_Explosion
Global Punkt() As Datentyp_Punkt

Sub DrawPoint(ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long)

  SetPixel HDC, X, Y, Color
  LastX = X
  LastY = Y

End Sub

Sub DrawLine(ByVal HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)

  Dim X As Long, Y As Long, I As Long
  Dim VarX As Double, VarY As Double
  Dim StepX As Double, StepY As Double
  Dim tmpX As Double, tmpY As Double
  Dim NegX As Long, NegY As Long
  Dim VAR As Long
  
  If ((LastX = X2) And (LastY = Y2)) Then Exit Sub
  
  VarX = Abs(X2 - X1)
  VarY = Abs(Y2 - Y1)
  
  If (VarX > VarY) Then
    VAR = VarX
    StepX = 1
    If (VarX = 0) Then StepY = 0 Else StepY = VarY / VarX
  Else
    VAR = VarY
    StepY = 1
    If (VarY = 0) Then StepX = 0 Else StepX = VarX / VarY
  End If

  If (X2 > X1) Then
    NegX = 1
  Else
    NegX = -1
  End If

  If (Y2 > Y1) Then
    NegY = 1
  Else
    NegY = -1
  End If

  For I = 0 To VAR
    tmpX = X1 + StepX * I * NegX
    tmpY = Y1 + StepY * I * NegY
    SetPixel HDC, tmpX, tmpY, Color
  Next
  
  LastX = X2
  LastY = Y2
  
End Sub

Sub PunkteInitialisieren()

  Dim I As Long
  
  For I = 1 To PunktCounter
    Punkt(I).Active = False
  Next

End Sub
