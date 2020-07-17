VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelCmd 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiles"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   6375
      Begin VB.TextBox HeightTxt 
         Height          =   285
         Left            =   5040
         TabIndex        =   11
         Text            =   "32"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox WidthTxt 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Text            =   "32"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alto del tile:"
         Height          =   195
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ancho del tile:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.CommandButton SearchInit 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox initPathTxt 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton SearchBmp 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox bmpPathTxt 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Accept 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Directorio Init:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Directorio con los bmp:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_HEIGHT As Byte = 32
Private Const DEFAULT_WIDTH As Byte = 32

Private Sub Accept_Click()
    If bmpPathTxt.Text = vbNullString Then
        MsgBox "Please select the BMP folder."
        Exit Sub
    End If
    
    If initPathTxt.Text = vbNullString Then
        MsgBox "Please select the Init folder."
        Exit Sub
    End If
    
    If Val(HeightTxt.Text) <= 0 Then
        MsgBox "Tile height must be a positive number and different from zero."
        Exit Sub
    End If
    
    If Val(WidthTxt.Text) <= 0 Then
        MsgBox "Tile width must be a positive number and different from zero."
        Exit Sub
    End If
    
    
    Config.bmpPath = bmpPathTxt.Text
    Config.initPath = initPathTxt.Text
    Config.TilePixelHeight = Val(HeightTxt.Text)
    Config.TilePixelWidth = Val(WidthTxt.Text)
    
    Config.SaveConfig
    
    Call Unload(Me)
End Sub

Private Sub CancelCmd_Click()
    'Is there a saved config, or we are requesting it for the first time?
    If Config.LoadConfig() Then
        Call Unload(Me)
    Else
        'Shut down!
        End
    End If
End Sub

Private Sub Form_Load()
    bmpPathTxt.Text = Config.bmpPath
    initPathTxt.Text = Config.initPath
    
    HeightTxt.Text = CStr(Config.TilePixelHeight)
    WidthTxt.Text = CStr(Config.TilePixelWidth)
End Sub

Private Sub HeightTxt_Change()
    If Not IsNumeric(HeightTxt.Text) Then
        HeightTxt.Text = DEFAULT_HEIGHT
    End If
End Sub

Private Sub SearchBmp_Click()
    bmpPathTxt.Text = BrowseForFolder(Me.hWnd, "Ubicación de los bmps")
End Sub

Private Sub SearchInit_Click()
    initPathTxt.Text = BrowseForFolder(Me.hWnd, "Ubicación de la carpeta Init")
End Sub

Private Sub WidthTxt_Change()
    If Not IsNumeric(WidthTxt.Text) Then
        WidthTxt.Text = DEFAULT_WIDTH
    End If
End Sub
