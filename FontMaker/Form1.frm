VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   3495
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4320
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar"
      Height          =   735
      Left            =   4200
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5280
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   5400
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   1
      Top             =   360
      Width           =   720
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3840
      Left            =   240
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   120
      Width           =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Pos
    X As Long
    Y As Long
    X2 As Long
    Y2 As Long
End Type
Dim Caracteres(255) As Pos

Private Sub Command1_Click()
Picture2.Cls
Picture2.PaintPicture Picture1.Image, 0, 0, Caracteres(Asc(Text1.Text)).X2, Caracteres(Asc(Text1.Text)).Y2, Caracteres(Asc(Text1.Text)).X, Caracteres(Asc(Text1.Text)).Y, Caracteres(Asc(Text1.Text)).X2, Caracteres(Asc(Text1.Text)).Y2
End Sub

Private Sub Command2_Click()
Dim st As String
Open (App.Path & "\font.ind") For Binary As #1
    For i = 1 To 255
        Put #1, , Caracteres(i)
        
        st = st & "char id=" & LPad(i, 5) & "x=" & LPad(Caracteres(i).X, 6) & "y=" & LPad(Caracteres(i).Y, 6) & "width=" & LPad(Caracteres(i).X2, 6) & "height=" & LPad(Caracteres(i).Y2, 6) & "xoffset=0     yoffset=0     xadvance=" & LPad(Caracteres(i).X2, 6) & "page=0  chnl=0 " & vbCrLf
    Next i
Close #1
Text2.Text = st
Call SavePicture(Picture1.Image, App.Path & "\fuente.bmp")
End Sub

Function LPad(ByVal s As String, cant As Integer) As String
Dim ss As String
ss = s
For i = 1 To cant - Len(s)
    ss = ss & " "
Next i
LPad = ss
End Function

Private Sub Form_Load()
Dim Texto As String
Dim Ax As Long
Dim Ay As Long
For i = 1 To 255
    Texto = Texto & Chr(i) & " " & IIf(Int(i / 25) = i / 25, vbCrLf, "")
    Picture1.ForeColor = RGB(10, 10, 10)
    Picture1.CurrentX = Ax
    Picture1.CurrentY = Ay
    Picture1.Print Chr(i) & " "
    Picture1.ForeColor = vbWhite
    Picture1.CurrentX = Ax + 2
    Picture1.CurrentY = Ay + 2
    Picture1.Print Chr(i) & " "
    With Caracteres(i)
        .X = Ax
        .Y = Ay
        .X2 = Picture1.TextWidth(Chr(i) & " ")
        .Y2 = Picture1.TextHeight(Chr(i) & " ")
    End With
    Ax = Ax + Picture1.TextWidth(Chr(i) & " ") + 2
    If Int(i / 18) = i / 18 Then Ax = 0: Ay = Ay + 16
Next i
Me.Caption = Caracteres(123).X
'Picture1.ForeColor = vbBlack
'Picture1.Print Texto
'Picture1.CurrentX = 2
'Picture1.CurrentY = 2

'Picture1.Print Texto
End Sub

