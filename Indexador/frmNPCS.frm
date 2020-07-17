VERSION 5.00
Begin VB.Form frmNPCS 
   Caption         =   "NPCS"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   17895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "E"
      Height          =   255
      Left            =   10680
      TabIndex        =   32
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Left            =   10440
      TabIndex        =   31
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   255
      Left            =   10200
      TabIndex        =   30
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox tCant 
      Height          =   285
      Left            =   9480
      TabIndex        =   29
      Text            =   "0"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox tObj 
      Height          =   285
      Left            =   7920
      TabIndex        =   28
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   7800
      TabIndex        =   27
      Top             =   5880
      Width           =   975
   End
   Begin VB.ListBox lInv 
      Height          =   4155
      Left            =   7920
      TabIndex        =   25
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox tFiltro 
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmNPCS.frx":0000
      Left            =   120
      List            =   "frmNPCS.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox tCampo 
      Height          =   285
      Index           =   9
      Left            =   4920
      TabIndex        =   21
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox tCampo 
      Height          =   285
      Index           =   8
      Left            =   4920
      TabIndex        =   19
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox tCampo 
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   17
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox tCampo 
      Height          =   285
      Index           =   6
      Left            =   4920
      TabIndex        =   15
      Top             =   4200
      Width           =   2775
   End
   Begin VB.TextBox tCampo 
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   13
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox tCampo 
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   11
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox tCampo 
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   9
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox tCampo 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   7
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox tCampo 
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox tCampo 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.PictureBox currentPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   6255
      Left            =   11520
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   2
      Top             =   120
      Width           =   6255
   End
   Begin VB.ListBox lNPC 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label lObj 
      Height          =   255
      Left            =   7920
      TabIndex        =   33
      Top             =   5160
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Inventario"
      Height          =   255
      Left            =   7920
      TabIndex        =   26
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label LCampo 
      Caption         =   "Evasión:"
      Height          =   255
      Index           =   9
      Left            =   4920
      TabIndex        =   22
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label LCampo 
      Caption         =   "Defensa:"
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   20
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label LCampo 
      Caption         =   "Ataque:"
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   18
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label LCampo 
      Caption         =   "Golpe:"
      Height          =   255
      Index           =   6
      Left            =   4920
      TabIndex        =   16
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label LCampo 
      Caption         =   "Vida:"
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   14
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label LCampo 
      Caption         =   "Oro:"
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label LCampo 
      Caption         =   "Experiencia:"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   10
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label LCampo 
      Caption         =   "Cabeza:"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label LCampo 
      Caption         =   "Cuerpo:"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label LCampo 
      Caption         =   "Nombre:"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Listado de NPC:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmNPCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Num As Integer
Private Sub Combo1_Click()
CargarLista (Combo1.ListIndex)
End Sub

Private Sub Command1_Click()
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "Name", tCampo(0).text)
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "Body", tCampo(1).text)
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "Head", tCampo(2).text)
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "GiveEXP", tCampo(3).text)
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "GiveGLD", tCampo(4).text)
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "MinHP", tCampo(5).text)
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "MaxHP", tCampo(5).text)
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "MinHIT", ReadField(1, tCampo(6).text, Asc("/")))
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "MaxHIT", ReadField(2, tCampo(6).text, Asc("/")))
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "PoderAtaque", tCampo(7).text)
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "DEF", tCampo(8).text)
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "PoderEvasion", tCampo(9).text)
NpcData(Num).name = tCampo(0).text
AbrirGrh (BodyData(Val(tCampo(1).text)).Walk(SOUTH).grhindex)
End Sub

Private Sub Command2_Click()

    Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "NROITEMS", lInv.ListCount + 1)
    

    Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "Obj" & lInv.ListCount + 1, tObj.text & "-" & tCant.text & " '" & lObj.Caption)
lInv.AddItem tObj.text & " - " & lObj.Caption & " - " & tCant.text

End Sub

Private Sub Command3_Click()
If lInv.ListIndex >= 0 Then
    Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "NROITEMS", lInv.ListCount - 1)
    
    For i = lInv.ListIndex + 1 To lInv.ListCount - 1
        Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "Obj" & i, GetVar(datPath & "NPCs.dat", "NPC" & Num, "Obj" & (i + 1)))
    Next i
    
    lInv.RemoveItem (lInv.ListIndex)

End If
End Sub

Private Sub Command4_Click()
If lInv.ListIndex < 0 Then Exit Sub
Dim i As Integer
i = lInv.ListIndex + 1
Call WriteVar(datPath & "NPCs.dat", "NPC" & Num, "Obj" & i, tObj.text & "-" & tCant.text & " '" & lObj.Caption)
lInv.List(i - 1) = tObj.text & " - " & lObj.Caption & " - " & tCant.text
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
End Sub

Sub CargarLista(Tipo As Integer)
Dim X As Integer
lNPC.Clear
For X = 1 To NumNPCs
            If NpcData(X).name <> "" And UCase(tFiltro.text) = UCase(Left(NpcData(X).name, Len(tFiltro.text))) Then
            
            Select Case Tipo
                Case 0
                    lNPC.AddItem X & " - " & NpcData(X).name
                Case 1
                    If NpcData(X).Comercia = 1 Then
                        lNPC.AddItem X & " - " & NpcData(X).name
                    End If
                Case 2
                    If NpcData(X).Alineacion = 2 And NpcData(X).NPCType <> 2 And X < 900 Then
                        lNPC.AddItem X & " - " & NpcData(X).name
                    End If
                Case 3
                    If NpcData(X).Alineacion = 0 And NpcData(X).Comercia = 0 And NpcData(X).NPCType <> 2 Then
                        lNPC.AddItem X & " - " & NpcData(X).name
                    End If
                Case 4
                    If NpcData(X).NPCType = 2 Then
                        lNPC.AddItem X & " - " & NpcData(X).name
                    End If
                Case 5
                    If X >= 900 Then
                        lNPC.AddItem X & " - " & NpcData(X).name
                    End If
            End Select
            End If
        Next X
End Sub

Private Sub lInv_Click()

If lInv.ListIndex < 0 Then Exit Sub

Dim Obj As String
Dim i As Integer
i = lInv.ListIndex + 1
Obj = GetVar(datPath & "NPCs.dat", "NPC" & Num, "Obj" & i)

tObj.text = Val(ReadField(1, Obj, 45))
tCant.text = Val(ReadField(2, Obj, 45))
End Sub

Private Sub lNPC_Click()
Dim NPC As Integer
    Dim Leer As New clsIniReader

    Call Leer.Initialize(datPath & "NPCs.dat")
NPC = ReadField(1, lNPC.text, 45)
Num = NPC
tCampo(0).text = Leer.GetValue("NPC" & NPC, "Name")
tCampo(1).text = Leer.GetValue("NPC" & NPC, "Body")
tCampo(2).text = Leer.GetValue("NPC" & NPC, "Head")
tCampo(3).text = Leer.GetValue("NPC" & NPC, "GiveEXP")
tCampo(4).text = Leer.GetValue("NPC" & NPC, "GiveGLD")
tCampo(5).text = Leer.GetValue("NPC" & NPC, "MaxHP")
tCampo(6).text = Leer.GetValue("NPC" & NPC, "MinHIT") & "/" & Leer.GetValue("NPC" & NPC, "MaxHIT")
tCampo(7).text = Leer.GetValue("NPC" & NPC, "PoderAtaque")
tCampo(8).text = Leer.GetValue("NPC" & NPC, "DEF")
tCampo(9).text = Leer.GetValue("NPC" & NPC, "PoderEvasion")
Dim NumInv As Integer
Dim i As Integer
Dim Obj As String
Dim ObjI As Integer
lInv.Clear
NumInv = Val(Leer.GetValue("NPC" & NPC, "NROITEMS"))
For i = 1 To NumInv
    Obj = Leer.GetValue("NPC" & NPC, "Obj" & i)
    If Obj <> "" Then
    ObjI = ReadField(1, Obj, 45)
    lInv.AddItem (ObjI & " - " & ObjData(ObjI).name & " - " & Val(ReadField(2, Obj, 45)))
    End If
Next i


Dim Cuerpo As Integer
Cuerpo = Val(tCampo(1).text)
Dim Cabeza As Integer
Cabeza = Val(tCampo(2).text)

AbrirGrh (BodyData(Cuerpo).Walk(SOUTH).grhindex)
End Sub

Private Sub tCampo_Change(Index As Integer)
If Val(tCampo(5).text) > 0 Then
LCampo(3).Caption = "Experiencia: Relacion: " & Round(Val(tCampo(3).text) / Val(tCampo(5).text), 3)
End If
End Sub

Private Sub tCant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command2_Click
End If
End Sub

Private Sub tFiltro_Change()
Combo1_Click
End Sub
Sub AbrirGrh(GI As Integer)
    
    Dim path As String

If GI = 0 Then
    previewer.Cls
    currentPic.Picture = Nothing
    Exit Sub
End If
FinCarga = False


    currentGrh = GI
    CurrentFrame = 1

    If GrhData(currentGrh).NumFrames > 0 Then
    If Right$(Config.bmpPath, 1) <> "\" Then
        path = Config.bmpPath & "\" & GrhData(GrhData(currentGrh).Frames(CurrentFrame)).FileNum & ".bmp"
    Else
        path = Config.bmpPath & GrhData(GrhData(currentGrh).Frames(CurrentFrame)).FileNum & ".bmp"
    End If
    End If
    'Prevent memory leaks
    Set currentPic.Picture = Nothing

    If Dir(path) <> "" Then
    Set currentPic.Picture = LoadPicture(path)
    End If

End Sub

Private Sub tObj_Change()
On Error Resume Next
lObj.Caption = ObjData(Val(tObj.text)).name
End Sub

Private Sub tObj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim F As New frmBuscarObj
    F.Show vbModal
    If F.NumObj > 0 Then
        tObj.text = F.NumObj
        tCant.SetFocus
        tCant.SelStart = 0
        tCant.SelLength = Len(tCant.text)
    End If
End If
End Sub
