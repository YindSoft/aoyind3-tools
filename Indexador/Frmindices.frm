VERSION 5.00
Begin VB.Form Frmindices 
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "salir"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.ListBox lstCual 
      Height          =   3960
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   7575
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Cuerpos"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Cabezas"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Cascos"
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Armas"
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Fx"
      Height          =   375
      Index           =   4
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Escudos"
      Height          =   375
      Index           =   5
      Left            =   5280
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Frmindices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Cant As Integer
lstCual.Clear
Cual = Index
Select Case Cual
    Case 0
        Cant = NumBodies
    Case 1
        Cant = Numheads
    Case 2
        Cant = NumCascos
    Case 3
        Cant = NumWeaponAnims
    Case 4
        Cant = NumFxs
    Case 5
        Cant = NumEscudosAnims
End Select
tNext.text = Cant + 1
Dim SinUso As Boolean
Dim TextoAd As String
Dim X As Integer
For i = 1 To Cant
SinUso = False
TextoAd = ""
Select Case Cual
    Case 0
        If BodyData(i).Walk(1).grhindex = 0 Then
            SinUso = True
        ElseIf GrhData(BodyData(i).Walk(1).grhindex).NumFrames = 0 Then
            SinUso = True
        End If
        For X = 1 To NumOBJs
            If ObjData(X).Ropaje = i And ObjData(X).ObjType <> 17 And ObjData(X).ObjType <> 16 And ObjData(X).ObjType <> 2 Then
                TextoAd = TextoAd & " - (" & X & ") " & ObjData(X).name
            End If
        Next X
        For X = 1 To NumNPCs
            If NpcData(X).Body = i Then
                TextoAd = TextoAd & " - (" & X & ") " & NpcData(X).name
            End If
        Next X
    Case 1
        If HeadData(i).Head(1).grhindex = 0 Then
            SinUso = True
        ElseIf GrhData(HeadData(i).Head(1).grhindex).NumFrames = 0 Then
            SinUso = True
        End If
    Case 2
        If CascoAnimData(i).Head(1).grhindex = 0 Then
            SinUso = True
        ElseIf GrhData(CascoAnimData(i).Head(1).grhindex).NumFrames = 0 Then
            SinUso = True
        End If
        For X = 1 To NumOBJs
            If ObjData(X).WeaponAnim = i And ObjData(X).ObjType = 17 Then
                TextoAd = TextoAd & " - (" & X & ") " & ObjData(X).name
            End If
        Next X
    Case 3
        If WeaponAnimData(i).WeaponWalk(1).grhindex = 0 Then
            SinUso = True
        ElseIf GrhData(WeaponAnimData(i).WeaponWalk(1).grhindex).NumFrames = 0 Then
            SinUso = True
        End If
        For X = 1 To NumOBJs
            If ObjData(X).WeaponAnim = i And ObjData(X).ObjType = 2 Then
                TextoAd = TextoAd & " - (" & X & ") " & ObjData(X).name
            End If
        Next X
    Case 4
        If FxData(i).Animacion = 0 Then
            SinUso = True
        ElseIf GrhData(FxData(i).Animacion).NumFrames = 0 Then
            SinUso = True
        End If
        'khalem
        For X = 1 To NumeroHechizos
            If Hechizos(X).FXgrh = i Then
                TextoAd = TextoAd & " - (" & X & ") " & Hechizos(X).Nombre
            End If
        Next X
    Case 5
        If ShieldAnimData(i).ShieldWalk(1).grhindex = 0 Then
            SinUso = True
        ElseIf GrhData(ShieldAnimData(i).ShieldWalk(1).grhindex).NumFrames = 0 Then
            SinUso = True
        End If
        For X = 1 To NumOBJs
            If ObjData(X).WeaponAnim = i And ObjData(X).ObjType = 16 Then
                TextoAd = TextoAd & " - (" & X & ") " & ObjData(X).name
            End If
        Next X
End Select
    If SinUso Then
        lstCual.AddItem i & " (Sin Uso) " & TextoAd
    Else
        lstCual.AddItem i & TextoAd
    End If
Next i

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub lstCual_Click()
Dim Index As Integer
Dim grhindex As Integer
Index = lstCual.ListIndex + 1
Select Case Cual
    Case 0
        grhindex = BodyData(Index).Walk(DirX).grhindex
        tOffX.text = BodyData(Index).HeadOffset.X
        tOffY.text = BodyData(Index).HeadOffset.Y
    Case 1
        grhindex = HeadData(Index).Head(DirX).grhindex
    Case 2
        grhindex = CascoAnimData(Index).Head(DirX).grhindex
    Case 3
        grhindex = WeaponAnimData(Index).WeaponWalk(DirX).grhindex
    Case 4
        grhindex = FxData(Index).Animacion
        tOffX.text = FxData(Index).OFFSETX
        tOffY.text = FxData(Index).OFFSETY
    Case 5
        grhindex = ShieldAnimData(Index).ShieldWalk(DirX).grhindex
End Select
    AbrirGrh (grhindex)
End Sub
