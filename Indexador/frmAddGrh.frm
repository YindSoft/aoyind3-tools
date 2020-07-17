VERSION 5.00
Begin VB.Form frmAddGrh 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Grh"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3180
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelCmd 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton OkCmd 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox indexTxt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox otherChk 
      Caption         =   "Usar otro"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1005
      Width           =   1215
   End
   Begin VB.ComboBox indexCmb 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Indices disponibles:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   1380
   End
End
Attribute VB_Name = "frmAddGrh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelCmd_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames = 0 Then
            Call indexCmb.AddItem(i)
        End If
    Next i
    
    'Add one after the last one
    Call indexCmb.AddItem(i)
    
    'Choose first one by default
    indexCmb.ListIndex = 0
End Sub

Private Sub indexTxt_Change()
    'Valdiate it's numeric
    If Not IsNumeric(indexTxt.Text) Then
        indexTxt.Text = Val(indexTxt.Text)
    End If
    
    'Prevent overflows
    If Val(indexTxt.Text) > &H7FFFFFFF Then
        indexTxt.Text = &H7FFFFFFF
    End If
    
    'Prevent undeflows
    If Val(indexTxt.Text) < 1 Then
        indexTxt.Text = "1"
    End If
    
    'Prevent the existance of decimals
    If InStr(1, indexTxt.Text, ",") Then
        indexTxt.Text = Left$(indexTxt.Text, InStr(1, indexTxt.Text, ",") - 1) & Mid$(indexTxt.Text, InStr(1, indexTxt.Text, ",") + 1)
    End If
    
    If InStr(1, indexTxt.Text, ".") Then
        indexTxt.Text = Left$(indexTxt.Text, InStr(1, indexTxt.Text, ".") - 1) & Mid$(indexTxt.Text, InStr(1, indexTxt.Text, ".") + 1)
    End If
End Sub

Private Sub OkCmd_Click()
    Dim index As Long
    Dim i As Long
    
    'Which index are we adding?
    If otherChk.value = vbChecked Then
        index = CLng(indexTxt.Text)
    Else
        index = CLng(indexCmb.Text)
    End If
    
    'Make sure he is not overwritting anything
    If index <= UBound(GrhData()) Then
        If GrhData(index).NumFrames > 0 Then
            If MsgBox("The chosen index is currently in use. Do you want to overwrite it?", vbOKCancel) = vbCancel Then
                Exit Sub
            End If
        End If
    Else
        'Resize array
        ReDim Preserve GrhData(1 To index) As GrhData
    End If
    
    If GrhData(index).NumFrames = 0 Then
        'Search where to place the grh....
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) > index Then
                Exit For
            End If
        Next i
        
        'Add it!
        Call frmMain.grhList.AddItem(index, i)
    Else
        'Search for the grh index within the grhList
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) = index Then
                Exit For
            End If
        Next i
    End If
    
    'Fill in grh data
    With GrhData(index)
        .FileNum = Val(frmMain.bmpTxt.Text)
        
        ReDim .Frames(1 To 1) As Integer
        .Frames(1) = index
        
        .NumFrames = 1
        .pixelHeight = Val(frmMain.grhHeightTxt.Text)
        .pixelWidth = Val(frmMain.grhWidthTxt.Text)
        .Speed = 0
        .sX = Val(frmMain.grhXTxt.Text)
        .sY = Val(frmMain.grhYTxt.Text)
        .TileHeight = .pixelHeight / Config.TilePixelHeight
        .TileWidth = .pixelWidth / Config.TilePixelWidth
    End With
    
    'Now select it in the list
    frmMain.grhList.ListIndex = i
    
    'Cya!
    Call Unload(Me)
End Sub

Private Sub otherChk_Click()
    indexTxt.Enabled = (otherChk.value = vbChecked)
    
    If indexTxt.Enabled Then
        indexTxt.BackColor = &H80000005
    Else
        indexTxt.BackColor = &H8000000F
    End If
End Sub
