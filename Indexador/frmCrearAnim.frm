VERSION 5.00
Begin VB.Form frmCrearAnim 
   Caption         =   "Crear Animacion"
   ClientHeight    =   10365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tSpeed 
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Text            =   "0"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Timer animation 
      Enabled         =   0   'False
      Left            =   120
      Top             =   6840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox tBMP 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Text            =   "0"
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox previewer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   1440
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   13
      Top             =   120
      Width           =   1440
   End
   Begin VB.PictureBox currentPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   1440
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   12
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox tRenglon 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Text            =   "0"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox tNF 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "0"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox tH 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "0"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox tW 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "0"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox tY 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "0"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox tX 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lPosMouse 
      Caption         =   "Posición Mouse"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "BMP:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Renglon:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Num Frames:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Largo:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Ancho:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmCrearAnim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentFrame As Integer

Private Sub animation_Timer()
    Dim path As String
    
    'If an animated grh is chosen, animate!

        If Val(tNF.text) > 1 Then
            'Move to next animation frame!
            CurrentFrame = CurrentFrame + 1
            
            If CurrentFrame > Val(tNF.text) Then
                CurrentFrame = 1
            End If
            
            DrawFrame (CurrentFrame)
        End If

End Sub

Sub DrawFrame(CF As Integer)
Dim X As Integer
Dim Y As Integer
Dim W As Integer
Dim H As Integer
W = Val(tW.text)
H = Val(tH.text)
Call ObtenerXY(X, Y, CF)

        previewer.Cls
        previewer.Width = W * 15
        previewer.Height = H * 15
        currentPic.Cls
        Call TransparentBlt(previewer.hDC, 0, 0, W, H, currentPic.hDC, X, Y, W, H, vbMagenta)
        currentPic.Line (X, Y)-(X + W, Y + H), , B

End Sub
Sub ObtenerXY(ByRef X As Integer, ByRef Y As Integer, CF As Integer)
Dim R As Integer
R = Val(tRenglon.text)
If R = 0 Then
    X = Val(tX.text) + (CF - 1) * Val(tW.text)
    Y = Val(tY.text)
Else
    X = Val(tX.text) + ((CF - 1) Mod R) * Val(tW.text)
    Y = Val(tY.text) + Int((CF - 1) / R) * Val(tH.text)
End If
End Sub
Sub DibujarIt()
Dim X As Integer
Dim Y As Integer
X = Val(tX.text)
Y = Val(tY.text)
currentPic.Cls
currentPic.Line (X, Y)-(X + Val(tW.text), Y + Val(tH.text)), , B
End Sub

Private Sub Command1_Click()
Dim Index As Integer
Dim GI As Integer
Dim Cant As Integer
Dim i As Integer
Dim Primero As Integer
Cant = Val(tNF.text)
Dim X As Integer
Dim Y As Integer
For i = 1 To Cant
    GI = frmMain.AgregarGrh(0)
    With GrhData(GI)
        .FileNum = Val(tBMP.text)
        .NumFrames = 1
        ReDim .Frames(1 To 1)
        .Frames(1) = GI
        Call ObtenerXY(X, Y, i)
        .sX = X
        .sY = Y
        .pixelWidth = Val(tW.text)
        .pixelHeight = Val(tH.text)
        .TileWidth = 32
        .TileHeight = 32
    End With
    If i = 1 Then Primero = GI
Next i
Index = frmMain.AgregarGrh(0, True)

With GrhData(Index)
    .NumFrames = Cant
    ReDim .Frames(1 To Cant)
    For i = 1 To Cant
        .Frames(i) = Primero + i - 1
    Next i
    .Speed = Val(tSpeed.text) * Cant
End With
End Sub

Private Sub currentPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lPosMouse.Caption = "X: " & X & "   -   Y: " & Y

End Sub

Private Sub tBMP_Change()
On Error Resume Next
Dim bmp As Integer
bmp = Val(tBMP.text)
If bmp > 0 Then
            Dim path As String
            If Right$(Config.bmpPath, 1) <> "\" Then
                path = Config.bmpPath & "\" & bmp & ".bmp"
            Else
                path = Config.bmpPath & bmp & ".bmp"
            End If
            
            currentPic.Picture = LoadPicture(path)
End If
DibujarIt
End Sub

Private Sub tH_Change()
DibujarIt
End Sub

Private Sub tNF_Change()
On Error Resume Next
animation.Interval = Round(Val(tSpeed.text))
animation.Enabled = True
End Sub

Private Sub tSpeed_Change()
animation.Interval = Round(Val(tSpeed.text))
animation.Enabled = True
End Sub

Private Sub tW_Change()
DibujarIt
End Sub

Private Sub tX_Change()
DibujarIt
End Sub

Private Sub tY_Change()
DibujarIt
End Sub
