VERSION 5.00
Begin VB.Form frmDraw 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Drawing Example"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   739
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton S_Help 
      Caption         =   "Help"
      Height          =   375
      Left            =   10200
      TabIndex        =   21
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox F_Fallen_Pixel 
      Height          =   285
      Left            =   10320
      TabIndex        =   20
      Text            =   "150"
      Top             =   3840
      Width           =   495
   End
   Begin VB.Frame O_Fallen 
      Caption         =   "Fallen"
      Height          =   1095
      Left            =   9240
      TabIndex        =   17
      Top             =   3240
      Width           =   1695
      Begin VB.OptionButton O_Fallen_Rand 
         Caption         =   "Rand unten"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1400
      End
      Begin VB.OptionButton O_Fallen_Begrenzt 
         Caption         =   "Pixel:"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame O_Farben 
      Caption         =   "Color"
      Height          =   1095
      Left            =   9240
      TabIndex        =   14
      Top             =   4440
      Width           =   1695
      Begin VB.OptionButton O_Farben_Einfarbig 
         Caption         =   "one-color"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton O_Farben_Wechsel 
         Caption         =   "changes"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame O_Art 
      Caption         =   "Kind"
      Height          =   1095
      Left            =   9240
      TabIndex        =   9
      Top             =   5640
      Width           =   1695
      Begin VB.OptionButton O_Art_Punkte 
         Caption         =   "Points"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton O_Art_Linien 
         Caption         =   "Lines"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox F_DemoAnzahl 
      Alignment       =   2  'Zentriert
      Height          =   285
      Left            =   10200
      TabIndex        =   8
      Text            =   "250"
      Top             =   6810
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton S_Demo 
      Caption         =   "Demo"
      Height          =   375
      Left            =   9240
      TabIndex        =   7
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Frame O_DrawModus 
      Caption         =   "Draw Modus"
      Height          =   1095
      Left            =   9240
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
      Begin VB.OptionButton O_Closed 
         Caption         =   "Closed"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton O_Points 
         Caption         =   "Points"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton S_Ende 
      Caption         =   "End"
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton S_Clear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7815
      Left            =   240
      MousePointer    =   2  'Kreuz
      ScaleHeight     =   519
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   575
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label T_Info_Maximale_Punkte 
      Caption         =   "Max.Pkt.:"
      Height          =   255
      Left            =   9240
      TabIndex        =   13
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label T_Info_DemoMode 
      Caption         =   "In Demo-Mode:"
      Height          =   255
      Left            =   9240
      TabIndex        =   12
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label F_Koordinaten 
      Caption         =   "Coordinates:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  Randomize

  LastX = 0
  LastY = 0

  O_Closed = True
  O_Art_Linien = True
  O_Farben_Wechsel = True
  O_Fallen_Begrenzt = True
  
  DemoMode = False
  Explosion.DoIt = False
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  LastX = 0
  LastY = 0

End Sub

Private Sub O_Fallen_Begrenzt_Click()

  Call O_Fallen_Rand_Click

End Sub

Private Sub O_Fallen_Rand_Click()

  If (O_Fallen_Rand) Then
    F_Fallen_Pixel.Enabled = False
  Else
    F_Fallen_Pixel.Enabled = True
  End If

End Sub

Private Sub Picture1_Click()
  
  If (DemoMode) Then
    'Left Mouseklick while Mouse is NOT moved = small explosion
    Explosion.DoIt = True
    Explosion.Great = False
  End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim I As Long, tmpColor As Long

  If (DemoMode) Then
    If ((Button And 2) = 2) Then
      'Right Mouseclick while Mouse is moved = BIG explosion (all free points)
      Explosion.DoIt = True
      Explosion.Great = True
    End If
    For I = 1 To PunktCounter
      If (Not Punkt(I).Active) Then
        Punkt(I).Active = True
        Punkt(I).X = X
        Punkt(I).Y = Y
        Punkt(I).LastX = X
        Punkt(I).LastY = Y
        Punkt(I).FallMax = Punkt(I).Y + F_Fallen_Pixel + Int(80 * Rnd - 40)
        Farbe I
        Punkt(I).StepX = Int(10 * Rnd + 2)
        If (Punkt(I).StepX > 6) Then Punkt(I).StepX = Punkt(I).StepX - 13
        Punkt(I).StepY = Int(4 * Rnd + 4) * (-1)
        Exit For
      End If
    Next
    If (Explosion.DoIt) Then
      Explosion.DoIt = False
      Explosion.Count = 0
      For I = 1 To PunktCounter
        If (Not Punkt(I).Active) Then
          Punkt(I).Active = True
          Punkt(I).X = X
          Punkt(I).Y = Y
          Punkt(I).LastX = X
          Punkt(I).LastY = Y
          Punkt(I).FallMax = Punkt(I).Y + F_Fallen_Pixel + Int(80 * Rnd - 40)
          Farbe I
          Punkt(I).StepX = Int(50 * Rnd - 25)
          Punkt(I).StepY = Int(10 * Rnd - 8)
          Explosion.Count = Explosion.Count + 1
          If ((Explosion.Count > 48) And (Explosion.Great = False)) Then
            Exit For
          End If
        End If
      Next
    End If
    Exit Sub
  End If
  
  F_Koordinaten = "Coordinates:" & vbCrLf & "X: " & X & "    Y: " & Y
  
  If (Button > 0) Then
    If ((Button And 1) = 1) Then tmpColor = vbCyan
    If ((Button And 2) = 2) Then tmpColor = vbRed
    If ((Button And 4) = 4) Then tmpColor = vbYellow
    
    If ((O_Points) Or ((LastX = 0) And (LastY = 0))) Then
      DrawPoint Picture1.HDC, X, Y, tmpColor
    Else
      DrawLine Picture1.HDC, LastX, LastY, X, Y, tmpColor
    End If
  Else
    LastX = 0
    LastY = 0
  End If

End Sub

Private Sub S_Clear_Click()

  Picture1 = LoadPicture
  LastX = 0
  LastY = 0

End Sub

Private Sub S_Demo_Click()

  If (DemoMode = False) Then
    If ((Me.F_DemoAnzahl = "") Or _
      (Val(Me.F_DemoAnzahl) < 1) Or _
      (Val(Me.F_DemoAnzahl) > 1000)) Then Exit Sub
    DemoMode = True
    S_Demo.Caption = "End Demo!"
    S_Clear.Enabled = False
    O_DrawModus.Enabled = False
    Timer1.Interval = 10
    Timer1.Enabled = True
    PunktCounter = Val(Me.F_DemoAnzahl)
    ReDim Punkt(PunktCounter)
    PunkteInitialisieren
  Else
    DemoMode = False
    S_Demo.Caption = "Demo"
    S_Clear.Enabled = True
    O_DrawModus.Enabled = True
    Timer1.Enabled = False
    Picture1 = LoadPicture
  End If
  
End Sub

Private Sub S_Ende_Click()

  Unload Me

End Sub

Private Sub S_Example_Click()

  Dim I As Long
  
  For I = 1 To 100
    SetPixel Picture1.HDC, I, I, vbWhite
  Next

End Sub

Private Sub S_Help_Click()

  Dim TX As String
  Dim R1 As String, R2 As String
  R1 = vbCrLf
  R2 = R1 & R1
  
  TX = "Drawing Example © 2001 by Stefan Ebert" & R2 _
    & "Program start in ""Draw Modus"", means you can " _
    & "draw with Your 3 MouseButtons. Change 'Draw Modus' " _
    & "Option Box to see how Windows handels the 'MouseMove' " _
    & "action." & R2 _
    & "Click ""Demo"" to activate the special graphical demo. " _
    & "Move the Mouse over the black picture field and see how " _
    & "it works. Use the 3 other option Boxes to change the outlook " _
    & "of the points. A Left MouseClick do an explosion with 50 " _
    & "points. Right MouseClick while moving makes an really BIG " _
    & "explosion. Every point will be used then." & R2 _
    & """Max.Pkt"" is the maximum points that the programm is allowed " _
    & "to use." & R2 _
    & "Please give response to ""st.ebert@t-online.de""      Many thnxx   ;-)"
  
  MsgBox TX, vbInformation, "About"

End Sub

Private Sub Timer1_Timer()

  Dim I As Long, FLAG As Long
  
  Picture1 = LoadPicture
  
  FLAG = 0
  For I = 1 To PunktCounter
    If (Punkt(I).Active) Then
      Punkt(I).X = Punkt(I).X + Punkt(I).StepX
      If ((Punkt(I).X > Picture1.ScaleWidth) Or (Punkt(I).X < 1)) Then
        Punkt(I).Active = False
      End If
      If (Punkt(I).Y > Picture1.ScaleHeight) Then
        Punkt(I).Active = False
      End If
      If ((Punkt(I).Y > Punkt(I).FallMax) And (O_Fallen_Begrenzt)) Then
        Punkt(I).Active = False
      End If
      If (Punkt(I).Active) Then
        FLAG = FLAG + 1
        Zeichnen (I) 'draw point or line
        Punkt(I).Y = Punkt(I).Y + Punkt(I).StepY * 2
        Punkt(I).StepY = Punkt(I).StepY + 1
      End If
    End If
  Next

  'boredom: make some example points !
  If (FLAG < 10) Then
    Picture1_MouseMove 0, 0, Int(Picture1.ScaleWidth * Rnd + 1), Picture1.ScaleHeight
  End If

  Me.F_Koordinaten.Caption = "Used points: " & FLAG
  DoEvents

End Sub

Sub Zeichnen(Nr As Long)

  If (O_Art_Linien) Then
    DrawLine Picture1.HDC, Punkt(Nr).LastX, Punkt(Nr).LastY, Punkt(Nr).X, Punkt(Nr).Y, Punkt(Nr).Color
    Punkt(Nr).LastX = Punkt(Nr).X
    Punkt(Nr).LastY = Punkt(Nr).Y
  Else
    DrawPoint Picture1.HDC, Punkt(Nr).X, Punkt(Nr).Y, Punkt(Nr).Color
  End If

End Sub

Sub Farbe(Nr As Long)

  If (O_Farben_Wechsel) Then
    Punkt(Nr).Color = QBColor(Int(15 * Rnd + 1))
  Else
    Punkt(Nr).Color = vbCyan
  End If

End Sub
