VERSION 5.00
Begin VB.Form frmBuscarObj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Obj"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   316
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lObj 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.TextBox tFiltro 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmBuscarObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NumObj As Integer

Private Sub Form_Load()
CargarLista
End Sub

Private Sub tFiltro_Change()
CargarLista
End Sub
Sub CargarLista()
Dim X As Integer
lObj.Clear
For X = 1 To NumOBJs
            If ObjData(X).name <> "" And UCase(tFiltro.text) = UCase(Left(ObjData(X).name, Len(tFiltro.text))) Then
            
                lObj.AddItem X & " - " & ObjData(X).name
End If
        Next X
If lObj.ListCount > 0 Then lObj.ListIndex = 0
End Sub

Private Sub tFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    If lObj.ListCount < lObj.ListIndex + 1 Then
        lObj.ListIndex = lObj.ListIndex + 1
    Else
        lObj.ListIndex = 0
    End If
ElseIf KeyCode = vbKeyUp Then
    If lObj.ListIndex > 0 Then
        lObj.ListIndex = lObj.ListIndex - 1
    Else
        lObj.ListIndex = lObj.ListCount - 1
    End If
End If
End Sub

Private Sub tFiltro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    NumObj = Val(ReadField(1, lObj.List(lObj.ListIndex), 45))
    Unload Me
ElseIf KeyAscii = 27 Then
    NumObj = 0
    Unload Me
End If
End Sub
