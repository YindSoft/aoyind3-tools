VERSION 5.00
Begin VB.Form frmCopiar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copiar Grhs"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox tBMP 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Text            =   "0"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox tHasta 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Text            =   "0"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox tDesde 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "0"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Cambiar bmp por:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta grh:"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Copiar desde el grh:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmCopiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim index As Long, index2 As Long
    Dim i As Long
    Dim a As Long, B As Long
    Dim f As Integer
    'Which index are we adding?
        index2 = Val(frmMain.grhList.List(frmMain.grhList.ListCount - 1)) + 1
        a = CLng(tDesde.Text)
        B = CLng(tHasta.Text) - CLng(tDesde.Text)
        
        ReDim Preserve GrhData(1 To index2 + B) As GrhData
        
    For x = a To CLng(tHasta.Text)
    
    index = index2 + x - a
    'Make sure he is not overwritting anything
    If index <= UBound(GrhData()) Then
        If GrhData(index).NumFrames > 0 Then
            If MsgBox("The chosen index is currently in use. Do you want to overwrite it?", vbOKCancel) = vbCancel Then
                Exit Sub
            End If
        End If
    Else
        'Resize array
        
    End If
    

    'Fill in grh data
    With GrhData(index)
        .FileNum = Val(frmMain.bmpTxt.Text)
        
        If GrhData(x).NumFrames > 1 Then
            ReDim .Frames(1 To GrhData(x).NumFrames) As Integer
            For f = 1 To GrhData(x).NumFrames
                .Frames(f) = GrhData(x).Frames(f) + (index - x)
            Next f
            .FileNum = GrhData(x).FileNum
            .NumFrames = GrhData(x).NumFrames
.Speed = GrhData(x).Speed

        Else
            ReDim .Frames(1 To 1) As Integer
            .Frames(1) = index
            .FileNum = Val(tBMP.Text)
            .NumFrames = 1
            .Speed = 1
        End If
        .pixelHeight = GrhData(x).pixelHeight
        .pixelWidth = GrhData(x).pixelWidth
        
        .sX = GrhData(x).sX
        .sY = GrhData(x).sY
        .TileHeight = GrhData(x).TileHeight
        .TileWidth = GrhData(x).TileWidth
        
    End With
    If GrhData(index).NumFrames = 1 Then
        'Search where to place the grh....
                For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) > index Then
                Exit For
            End If
        Next i
        'Add it!
        Call frmMain.grhList.AddItem(index, i)
    Else
            For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) > index Then
                Exit For
            End If
        Next i
        'Search for the grh index within the grhList
        Call frmMain.grhList.AddItem(index & " (ANIMACIÓN)", i)
    End If
    
    Next x
    'Now select it in the list
    frmMain.grhList.ListIndex = i
    
    'Cya!
    Call Unload(Me)
End Sub

