VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indexador AoYind"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   ScaleHeight     =   709
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   Begin VB.CommandButton Command27 
      Caption         =   "NPC"
      Height          =   255
      Left            =   13440
      TabIndex        =   112
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Command25"
      Height          =   615
      Left            =   9720
      TabIndex        =   108
      Top             =   9960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Commanda 
      Caption         =   ">"
      Height          =   315
      Index           =   2
      Left            =   13800
      TabIndex        =   106
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Commanda 
      Caption         =   "^"
      Height          =   195
      Index           =   1
      Left            =   13560
      TabIndex        =   105
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Commanda 
      Caption         =   "^"
      Height          =   195
      Index           =   3
      Left            =   13560
      TabIndex        =   107
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Commanda 
      Caption         =   "<"
      Height          =   315
      Index           =   0
      Left            =   13440
      TabIndex        =   104
      Top             =   3720
      Width           =   255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Terreno"
      Height          =   1215
      Left            =   3000
      TabIndex        =   88
      Top             =   960
      Width           =   8895
      Begin VB.CommandButton Command26 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   1920
         TabIndex        =   111
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox tITSolo 
         Height          =   285
         Left            =   960
         TabIndex        =   110
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Indexar Terreno"
         Height          =   735
         Left            =   8040
         TabIndex        =   96
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox tITNombre 
         Height          =   285
         Left            =   5640
         TabIndex        =   95
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox tITX 
         Height          =   285
         Left            =   120
         TabIndex        =   94
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox tITY 
         Height          =   285
         Left            =   960
         TabIndex        =   93
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox tITAncho 
         Height          =   285
         Left            =   1800
         TabIndex        =   92
         Text            =   "2"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox tITLargo 
         Height          =   285
         Left            =   2640
         TabIndex        =   91
         Text            =   "2"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox tITbmp 
         Height          =   285
         Left            =   3480
         TabIndex        =   90
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox tITDesde 
         Height          =   285
         Left            =   4560
         TabIndex        =   89
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Solo el grh:"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Nombre Terreno"
         Height          =   255
         Left            =   5640
         TabIndex        =   103
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label23 
         Caption         =   "Y:"
         Height          =   255
         Left            =   960
         TabIndex        =   101
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Ancho:"
         Height          =   255
         Left            =   1800
         TabIndex        =   100
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "Largo:"
         Height          =   255
         Left            =   2640
         TabIndex        =   99
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label26 
         Caption         =   "BMP:"
         Height          =   255
         Left            =   3480
         TabIndex        =   98
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Desde Grh:"
         Height          =   255
         Left            =   4560
         TabIndex        =   97
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Zoom"
      Height          =   735
      Left            =   9480
      TabIndex        =   79
      Top             =   120
      Width           =   2055
      Begin VB.TextBox ZoomTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   82
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton ZoomIn 
         Caption         =   "+"
         Height          =   255
         Left            =   720
         TabIndex        =   81
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton ZoomOut 
         Caption         =   "-"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   1800
         TabIndex        =   83
         Top             =   285
         Width           =   120
      End
   End
   Begin VB.Frame grhFrame 
      Caption         =   "Grh"
      Height          =   735
      Left            =   3000
      TabIndex        =   68
      Top             =   120
      Width           =   6375
      Begin VB.TextBox grhXTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   360
         TabIndex        =   73
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhYTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   72
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhHeightTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         TabIndex        =   71
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox grhWidthTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   70
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox bmpTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5520
         TabIndex        =   69
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   1320
         TabIndex        =   77
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alto:"
         Height          =   195
         Left            =   3840
         TabIndex        =   76
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ancho:"
         Height          =   195
         Left            =   2400
         TabIndex        =   75
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bmp:"
         Height          =   195
         Left            =   5040
         TabIndex        =   74
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Crear Animacion"
      Height          =   255
      Left            =   7800
      TabIndex        =   67
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox tAnimDesde 
      Height          =   285
      Left            =   3000
      TabIndex        =   66
      Text            =   "0"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox tAnimHasta 
      Height          =   285
      Left            =   4200
      TabIndex        =   65
      Text            =   "0"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox tEnGrh 
      Height          =   285
      Left            =   5400
      TabIndex        =   64
      Text            =   "0"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6600
      TabIndex        =   63
      Text            =   "50"
      Top             =   2640
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   3960
      TabIndex        =   26
      Top             =   7200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buscador"
      Height          =   3855
      Left            =   0
      TabIndex        =   52
      Top             =   6600
      Width           =   2895
      Begin VB.CommandButton Command14 
         Caption         =   "Buscar BMP"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Graficos No Indexados"
         Height          =   495
         Left            =   1440
         TabIndex        =   59
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox tBuscar 
         Height          =   285
         Left            =   240
         TabIndex        =   58
         Text            =   "0"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Borrar Todo"
         Height          =   495
         Left            =   1800
         TabIndex        =   57
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Borrar Grh y BMP"
         Height          =   495
         Left            =   960
         TabIndex        =   56
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Borrar Grh"
         Height          =   495
         Left            =   120
         TabIndex        =   55
         Top             =   3240
         Width           =   855
      End
      Begin VB.ListBox lstBuscar 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CheckBox grhOnly 
         Caption         =   "Mostrar solamente el Grh"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Buscar Grh"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Recargar Dats"
      Height          =   375
      Left            =   12000
      TabIndex        =   48
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtAnim 
      Height          =   285
      Index           =   3
      Left            =   6600
      TabIndex        =   46
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtAnim 
      Height          =   285
      Index           =   2
      Left            =   5400
      TabIndex        =   45
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtAnim 
      Height          =   285
      Index           =   1
      Left            =   4200
      TabIndex        =   44
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtAnim 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   39
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Agregado Manual"
      Height          =   255
      Left            =   7800
      TabIndex        =   38
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox tOffY 
      Height          =   285
      Left            =   13440
      TabIndex        =   36
      Text            =   "0"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox tOffX 
      Height          =   285
      Left            =   13440
      TabIndex        =   34
      Text            =   "0"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox tDesdeGrh 
      Height          =   285
      Left            =   12120
      TabIndex        =   32
      Text            =   "0"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox tNBMP 
      Height          =   285
      Left            =   12120
      TabIndex        =   30
      Text            =   "0"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox tNext 
      Height          =   285
      Left            =   12120
      TabIndex        =   28
      Text            =   "0"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Crear Nuevo desde seleccionado"
      Height          =   495
      Left            =   12120
      TabIndex        =   27
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   7560
      TabIndex        =   22
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "0"
      Height          =   375
      Left            =   7200
      TabIndex        =   23
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   6840
      TabIndex        =   21
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Anch -1"
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Anch +1"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Derecha"
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Izq"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Alto -1"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Alto +1"
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subir"
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bajar"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox currentPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   9240
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   12
      Top             =   4800
      Width           =   4935
   End
   Begin VB.OptionButton optDir 
      Caption         =   "Izquierda"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   11280
      TabIndex        =   11
      Top             =   4200
      Width           =   1335
   End
   Begin VB.OptionButton optDir 
      Caption         =   "Abajo"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   11280
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
   End
   Begin VB.OptionButton optDir 
      Caption         =   "Derecha"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   10200
      TabIndex        =   9
      Top             =   4200
      Width           =   1095
   End
   Begin VB.OptionButton optDir 
      Caption         =   "Arriba"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   10200
      TabIndex        =   8
      Top             =   3960
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Escudos"
      Height          =   375
      Index           =   5
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Fx"
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Armas"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Cascos"
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Cabezas"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton btnCual 
      Caption         =   "Cuerpos"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Timer animation 
      Enabled         =   0   'False
      Left            =   120
      Top             =   1200
   End
   Begin VB.ListBox grhList 
      Height          =   5130
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0002
      TabIndex        =   0
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Quitar Grh Conocidos"
      Height          =   495
      Left            =   12000
      TabIndex        =   25
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Todos"
      Height          =   375
      Left            =   960
      TabIndex        =   51
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox lstCual 
      Height          =   5130
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.PictureBox previewer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   3000
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   416
      TabIndex        =   62
      Top             =   4800
      Width           =   6240
   End
   Begin VB.Label Label13 
      Caption         =   "Desde:"
      Height          =   255
      Left            =   3000
      TabIndex        =   87
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Hasta:"
      Height          =   255
      Left            =   4200
      TabIndex        =   86
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "En Grh:"
      Height          =   255
      Left            =   5400
      TabIndex        =   85
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   6600
      TabIndex        =   84
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label28 
      Caption         =   "GRHs:"
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lPosMouse 
      Caption         =   "Posición Mouse"
      Height          =   255
      Left            =   6600
      TabIndex        =   49
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lSpeed 
      Caption         =   "Speed: 0"
      Height          =   255
      Left            =   7800
      TabIndex        =   47
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label18 
      Caption         =   "Anim 4:"
      Height          =   255
      Left            =   6600
      TabIndex        =   43
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "Anim 3:"
      Height          =   255
      Left            =   5400
      TabIndex        =   42
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "Anim 2:"
      Height          =   255
      Left            =   4200
      TabIndex        =   41
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "Anim 1:"
      Height          =   255
      Left            =   3000
      TabIndex        =   40
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Offset Y:"
      Height          =   255
      Left            =   13440
      TabIndex        =   37
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Offset X:"
      Height          =   255
      Left            =   13440
      TabIndex        =   35
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Desde Grh:"
      Height          =   255
      Left            =   12120
      TabIndex        =   33
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "BMP:"
      Height          =   255
      Left            =   12120
      TabIndex        =   31
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Next Index:"
      Height          =   255
      Left            =   12120
      TabIndex        =   29
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Frame: 1"
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&File"
      Begin VB.Menu SaveMnu 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveOldMnu 
         Caption         =   "Save in &old format"
      End
      Begin VB.Menu SaveNewMnu 
         Caption         =   "Save in &new format"
      End
   End
   Begin VB.Menu GrhMnu 
      Caption         =   "&Grh"
      Begin VB.Menu AddGrhMnu 
         Caption         =   "&Agregar Grh..."
         Shortcut        =   ^N
      End
      Begin VB.Menu RemoveGrhMnu 
         Caption         =   "&Remover Grh"
         Shortcut        =   ^D
      End
      Begin VB.Menu mCopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu MMCB 
         Caption         =   "Agregar Grhs desde Clipboard"
      End
      Begin VB.Menu mCrearA 
         Caption         =   "Crear Animacion"
      End
   End
   Begin VB.Menu mAddgrh 
      Caption         =   "Agregar Grh Rapido"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Posiciones x y para utilizarn PaintPicture
Private Pos_x As Long
Private Pos_y As Long

'variables para almacenar el ancho y Alto de la imagen a cargar _
 en el control PictureBox
Private Ancho_Pic  As Single
Private Alto_Pic As Single

'Indica si estamos pulsando el mouse en el Command
Private Mouse_Abajo As Boolean
'Para variar el desplazamiento
Private Avance As Single

''
' Default zoom, 100%
Private Const DEFAULT_ZOOM As Integer = 100

''
' Maximum zoom possible, 10 times bigger.
Private Const MAX_ZOOM As Integer = DEFAULT_ZOOM * 10

''
' Minimum zoom possible, 10 times smaller.
Private Const MIN_ZOOM As Integer = DEFAULT_ZOOM / 10

''
' Step by which zoom is altered.
Private Const ZOOM_STEP As Integer = 10

''
' Means no grh is being rendered.
Private Const NO_GRH As Long = -1


''
' Defines the different points of the selection box that are being edited.
'
' @param    sbpeNone            No coord is being modified.
' @param    sbpeStartX          Starting x coord is being modified.
' @param    sbpeStartY          Starting y coord is being modified.
' @param    sbpeEndX            Ending x coord is being modified.
' @param    sbpeEndY            Ending y coord is being modified.
' @param    sbpeStartXStartY    Starting x coord and starting y coord are being modified.
' @param    sbpeEndXEndY        Ending x coord and ending y coord are being modified.
' @param    sbpeStartXEndY      Starting x coord and ending y coord are being modified.
' @param    sbpeEndXStartY      Ending x coord and starting y coord are being modified.

Private Enum eSelectionBoxPointEdition
    sbpeNone
    sbpeStartX
    sbpeStartY
    sbpeEndX
    sbpeEndY
    sbpeStartXStartY
    sbpeEndXEndY
    sbpeStartXEndY
    sbpeEndXStartY
End Enum

''
' The current zoom, 1 == 100%
Private zoom As Single

''
'Currently loaded picture. Used to render avoiding to reload everytime zoom or scroll happens.

''
' X coord where a selection started.
Private selectionAreaStartX As Single

''
' Y coord where a selection started.
Private selectionAreaStartY As Single

''
' X coord where a selection ended.
Private selectionAreaEndX As Single

''
' Y coord where a selection ended.
Private selectionAreaEndY As Single

''
' Cord currently being edited.
Private editionCoord As eSelectionBoxPointEdition

''
' The grh currently being displayed
Private currentGrh As Long

''
' The current frame of the grh being displayed
Private CurrentFrame As Long

''
' Flag used to ignore calls to RenderSelectionBox.
Private ignoreSelectionBoxRender As Boolean

''
' Flag used to ignore update events to grh' data textboxes.
Private ignoreGrhTextUpdate As Boolean


Dim FinCarga As Boolean
Dim Cual As Integer
Dim DirX As Integer
Private Sub AddGrhMnu_Click()
    Call frmAddGrh.Show(vbModal, Me)
End Sub

Private Sub animation_Timer()
    Dim path As String
    
    'If an animated grh is chosen, animate!
    If currentGrh <> NO_GRH Then
        If GrhData(currentGrh).NumFrames > 1 Then
            'Move to next animation frame!
            CurrentFrame = CurrentFrame + 1
            
            If CurrentFrame > GrhData(currentGrh).NumFrames Then
                CurrentFrame = 1
            End If
            
            DrawFrame (CurrentFrame)
        End If
    End If
End Sub
Sub DrawFrame(FrameIndex As Integer)
            'Load new bitmap
            On Error Resume Next
            Dim path As String
            If Right$(Config.bmpPath, 1) <> "\" Then
                path = Config.bmpPath & "\" & GrhData(GrhData(currentGrh).Frames(FrameIndex)).FileNum & ".bmp"
            Else
                path = Config.bmpPath & GrhData(GrhData(currentGrh).Frames(FrameIndex)).FileNum & ".bmp"
            End If
            
            'Prevent memory leaks
            Set currentPic.Picture = Nothing
            If Dir(path) <> "" Then
            Set currentPic.Picture = LoadPicture(path)
            End If
            Call RedrawPicture(currentGrh, FrameIndex)
            Label7.Caption = "Frame: " & FrameIndex & " / " & GrhData(currentGrh).NumFrames
Dim GI As Integer
GI = GrhData(currentGrh).Frames(FrameIndex)

grhFrame.Enabled = True
grhXTxt.text = GrhData(GI).sX
grhYTxt.text = GrhData(GI).sY
grhWidthTxt.text = GrhData(GI).pixelWidth
grhHeightTxt.text = GrhData(GI).pixelHeight
bmpTxt.text = GrhData(GI).FileNum
End Sub
Private Sub bmpTxt_Change()

If Not FinCarga Then Exit Sub
    Dim path As String
    
    'Prevent non numeric characters
    If Not IsNumeric(bmpTxt.text) Then
        bmpTxt.text = Val(bmpTxt.text)
    End If
    
    'Prevent overflow
    If Val(bmpTxt.text) > &H7FFFFFFF Then
        bmpTxt.text = &H7FFFFFFF
    End If
    
    'Prevent underrflow
    If Val(bmpTxt.text) < 1 Then
        bmpTxt.text = "1"
    End If
    
    
    If Right$(Config.bmpPath, 1) <> "\" Then
        path = Config.bmpPath & "\" & bmpTxt.text & ".bmp"
    Else
        path = Config.bmpPath & bmpTxt.text & ".bmp"
    End If
    
    'If file exists, load it
    If FileExists(path) And currentGrh <> NO_GRH Then
        GrhData(currentGrh).FileNum = CLng(bmpTxt.text)
        
        'Prevent memory leaks
        Set currentPic.Picture = Nothing
        Set currentPic.Picture = LoadPicture(path)
        
        'Set scrollers!
        Call SetScrollers
        
        'Display the grh!
        Call RedrawPicture(currentGrh, CurrentFrame)
        
        'Show selection box (if needed)
        ignoreSelectionBoxRender = (grhOnly.value = vbChecked)
        Call RenderSelectionBox
    End If
End Sub

Private Sub btnCual_Click(Index As Integer)
lstCual.Visible = True
grhList.Visible = False

Dim i As Integer
For i = 0 To 3
optDir(i).Enabled = True
Next i
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
End Sub

Private Sub Command1_Click()
Dim grhindex As Integer

        If GrhData(currentGrh).NumFrames > 1 Then

            grhindex = GrhData(currentGrh).Frames(CurrentFrame)
        Else
            grhindex = currentGrh
        End If

GrhData(grhindex).sY = GrhData(grhindex).sY - 1
DrawFrame (CurrentFrame)
End Sub

Private Sub Command10_Click()
        If GrhData(currentGrh).NumFrames > 1 Then
            'Move to next animation frame!
            CurrentFrame = CurrentFrame + 1
            
            If CurrentFrame > GrhData(currentGrh).NumFrames Then
                CurrentFrame = 1
            End If
            
            DrawFrame (CurrentFrame)
        End If
End Sub

Private Sub Command11_Click()
animation.Enabled = False
        If GrhData(currentGrh).NumFrames > 1 Then
            'Move to next animation frame!
            CurrentFrame = 1
        
            
            DrawFrame (CurrentFrame)
        End If
End Sub

Sub QuitarDeLista(grhindex As Integer)
Dim i As Integer
For i = 0 To grhList.ListCount - 1
    If Val(grhList.List(i)) = grhindex Then
        grhList.RemoveItem (i)
        Exit For
    End If
Next i
End Sub
Function EstaEnUso(grhindex As Integer)
Dim i As Integer
Dim X As Integer
Dim F As Integer
Dim GrhI As Integer
For i = 1 To NumBodies
    For X = 1 To 4
        GrhI = BodyData(i).Walk(X).grhindex
        If grhindex = GrhI Then
            EstaEnUso = True
            Exit Function
        End If
        For F = 1 To GrhData(GrhI).NumFrames
            If GrhData(GrhI).Frames(F) = grhindex Then
                EstaEnUso = True
                Exit Function
            End If
        Next F
    Next X
Next i
For i = 1 To NumWeaponAnims
    For X = 1 To 4
        GrhI = WeaponAnimData(i).WeaponWalk(X).grhindex
        If GrhI > 0 Then
        If grhindex = GrhI Then
            EstaEnUso = True
            Exit Function
        End If
        
        For F = 1 To GrhData(GrhI).NumFrames
            If GrhData(GrhI).Frames(F) = grhindex Then
                EstaEnUso = True
                Exit Function
            End If
        Next F
        End If
    Next X
Next i
For i = 1 To Numheads
    For X = 1 To 4
        GrhI = HeadData(i).Head(X).grhindex
        If GrhI > 0 Then
        If grhindex = GrhI Then
            EstaEnUso = True
            Exit Function
        End If
        For F = 1 To GrhData(GrhI).NumFrames
            If GrhData(GrhI).Frames(F) = grhindex Then
                EstaEnUso = True
                Exit Function
            End If
        Next F
        End If
    Next X
Next i
For i = 1 To NumCascos
    For X = 1 To 4
        GrhI = CascoAnimData(i).Head(X).grhindex
        If GrhI > 0 Then
        If grhindex = GrhI Then
            EstaEnUso = True
            Exit Function
        End If
        For F = 1 To GrhData(GrhI).NumFrames
            If GrhData(GrhI).Frames(F) = grhindex Then
                EstaEnUso = True
                Exit Function
            End If
        Next F
        End If
    Next X
Next i
For i = 1 To NumFxs
        GrhI = FxData(i).Animacion
        If grhindex = GrhI Then
            EstaEnUso = True
            Exit Function
        End If
        For F = 1 To GrhData(GrhI).NumFrames
            If GrhData(GrhI).Frames(F) = grhindex Then
                EstaEnUso = True
                Exit Function
            End If
        Next F
Next i
For i = 1 To NumEscudosAnims
    For X = 1 To 4
        GrhI = ShieldAnimData(i).ShieldWalk(X).grhindex
        If GrhI > 0 Then
        If grhindex = GrhI Then
            EstaEnUso = True
            Exit Function
        End If
        For F = 1 To GrhData(GrhI).NumFrames
            If GrhData(GrhI).Frames(F) = grhindex Then
                EstaEnUso = True
                Exit Function
            End If
        Next F
        End If
    Next X
Next i
End Function

Private Sub Command12_Click()
CargarLista (True)
End Sub

Private Sub Command13_Click()
Dim grhindex As Integer
Dim i As Integer
grhindex = Val(tBuscar.text)
'khalem
lstBuscar.Enabled = True
lstBuscar.Clear
If GrhData(grhindex).NumFrames = 1 Then
    lstBuscar.AddItem grhindex
Else
    lstBuscar.AddItem grhindex & " (ANIMACIÓN)"
    For i = 1 To GrhData(grhindex).NumFrames
        lstBuscar.AddItem GrhData(grhindex).Frames(i)
    Next i
End If
End Sub

Private Sub Command14_Click()
Dim grhindex As Integer
Dim i As Integer
Dim X As Integer
Dim Agrego As Boolean
grhindex = Val(tBuscar.text)
lstBuscar.Clear
'khalem
lstBuscar.Enabled = True

For i = 1 To grhCount
    If GrhData(i).FileNum = grhindex Then
        lstBuscar.AddItem i
    Else
        Agrego = False
        For X = 1 To GrhData(i).NumFrames
            If GrhData(GrhData(i).Frames(X)).FileNum = grhindex Then
                Agrego = True
                lstBuscar.AddItem GrhData(i).Frames(X)
            End If
        Next X
        If Agrego Then
            lstBuscar.AddItem i & " (ANIMACIÓN)"
        End If
    End If
Next i
End Sub

Function BMPenGRH(Num As Integer) As Integer
Dim i As Integer, X As Integer
For i = 1 To grhCount
    If GrhData(i).FileNum = Num Then
        BMPenGRH = i
        Exit Function
    Else
        For X = 1 To GrhData(i).NumFrames
            If GrhData(GrhData(i).Frames(X)).FileNum = Num Then
                BMPenGRH = i
                Exit Function
            End If
        Next X
    End If
Next i
BMPenGRH = 0
End Function

Private Sub Command15_Click()
Dim FSO As New FileSystemObject
Dim X As file
Dim Num As Integer
Dim F As Folder
Dim i As Integer

Set F = FSO.GetFolder(Config.bmpPath)
PB.max = F.Files.Count
For Each X In F.Files
    Num = Val(X.name)
    PB.value = i
    i = i + 1
    If Num > 0 Then
        If BMPenGRH(Num) = 0 Then
            FSO.MoveFile F.path & "\" & Num & ".bmp", App.path & "\NoUsados\" & Num & ".bmp"
        End If
    End If
Next X
End Sub

Sub BorrarGrh(GH As Integer, Optional ByVal BorrarBMP As Boolean = True)
Dim bmp As Integer
Dim i As Integer
If GH > 0 Then
bmp = GrhData(GH).FileNum
If BorrarBMP Then
    If Dir(Config.bmpPath & "\" & bmp & ".bmp") <> "" Then
        Kill (Config.bmpPath & "\" & bmp & ".bmp")
    End If
End If
        With GrhData(GH)
            .FileNum = 0
            .NumFrames = 0
            .Frames(1) = 0
            .pixelHeight = 0
            .pixelWidth = 0
            .sX = 0
            .sY = 0
            .Speed = 0
            .TileHeight = 0
            .TileWidth = 0
        End With
For i = 0 To grhList.ListCount - 1
    If Val(grhList.List(i)) = GH Then
        grhList.RemoveItem i
        Exit For
    End If
Next i
For i = 0 To lstBuscar.ListCount - 1
    If Val(lstBuscar.List(i)) = GH Then
        lstBuscar.RemoveItem i
        Exit For
    End If
Next i
End If
End Sub

Private Sub Command16_Click()
Dim i As Integer
Dim GI As Integer
Dim cG As Integer
cG = Val(lstBuscar.text)
'If MsgBox("¿Esta seguro que desea borrar la indexacion y el bmp correspondiente al grh " & cG & "?", vbExclamation + vbYesNo) = vbYes Then
    If GrhData(currentGrh).NumFrames > 0 Then
        For i = 1 To GrhData(cG).NumFrames
            GI = GrhData(cG).Frames(i)
            BorrarGrh (GI)
        Next i
        BorrarGrh (cG)
    Else
        BorrarGrh (cG)
    End If
'End If
End Sub

Private Sub Command17_Click()
Dim i As Integer
Dim GI As Integer
Dim cG As Integer
cG = Val(lstBuscar.text)
'If MsgBox("¿Esta seguro que desea borrar la indexacion y el bmp correspondiente al grh " & cG & "?", vbExclamation + vbYesNo) = vbYes Then
    If GrhData(currentGrh).NumFrames > 0 Then
        For i = 1 To GrhData(cG).NumFrames
            GI = GrhData(cG).Frames(i)
            Call BorrarGrh(GI, False)
        Next i
        Call BorrarGrh(cG, False)
    Else
        Call BorrarGrh(cG, False)
    End If
'End If
End Sub

Private Sub Command18_Click()
Dim N As Integer
Dim B As Integer
Dim i As Integer
Dim DGH As Integer
N = Val(tNext.text)
B = Val(lstCual.text)
If B = 0 Or N = 0 Then Exit Sub
DGH = Val(tDesdeGrh.text)
Select Case Cual
    Case 0
        If N > NumBodies Then
            NumBodies = NumBodies + 1
            ReDim Preserve MisCuerpos(1 To NumBodies)
            ReDim Preserve BodyData(1 To NumBodies)
        End If
        MisCuerpos(N).HeadOffsetX = MisCuerpos(B).HeadOffsetX
        MisCuerpos(N).HeadOffsetY = MisCuerpos(B).HeadOffsetY
        BodyData(N).HeadOffset.X = MisCuerpos(N).HeadOffsetX
        BodyData(N).HeadOffset.Y = MisCuerpos(N).HeadOffsetY
        For i = 1 To 4
            MisCuerpos(N).Body(i) = CopiarAnimacion(MisCuerpos(B).Body(i), Val(tNBMP.text), DGH)
            InitGrh BodyData(N).Walk(i), MisCuerpos(N).Body(i), 0
        Next i

    Case 1
        If N > Numheads Then
            Numheads = Numheads + 1
            ReDim Preserve HeadData(1 To Numheads)
            ReDim Preserve MisCabezas(1 To Numheads)
        End If
        For i = 1 To 4
            MisCabezas(N).Head(i) = CopiarAnimacion(MisCabezas(B).Head(i), Val(tNBMP.text), DGH)
            InitGrh HeadData(N).Head(i), MisCabezas(N).Head(i), 0
        Next i
    Case 2
        If N > NumCascos Then
            NumCascos = NumCascos + 1
            ReDim Preserve MisCascos(1 To NumCascos) As tIndiceCabeza
            ReDim Preserve CascoAnimData(1 To NumCascos) As HeadData
        End If
        For i = 1 To 4
            MisCascos(N).Head(i) = CopiarAnimacion(MisCascos(B).Head(i), Val(tNBMP.text), DGH)
            InitGrh CascoAnimData(N).Head(i), MisCascos(N).Head(i), 0
        Next i
    Case 3
        If N > NumWeaponAnims Then
            NumWeaponAnims = NumWeaponAnims + 1
            ReDim Preserve MisArmas(1 To NumWeaponAnims)
            ReDim Preserve WeaponAnimData(1 To NumWeaponAnims)
        End If
        For i = 1 To 4
            MisArmas(N).Arma(i) = CopiarAnimacion(MisArmas(B).Arma(i), Val(tNBMP.text), DGH)
            InitGrh WeaponAnimData(N).WeaponWalk(i), MisArmas(N).Arma(i), 0
        Next i
    Case 4
        If N > NumFxs Then
            NumFxs = NumFxs + 1
            ReDim Preserve FxData(1 To NumFxs)
        End If
        FxData(N).OFFSETX = FxData(B).OFFSETX
        FxData(N).OFFSETY = FxData(B).OFFSETY
        FxData(N).Animacion = CopiarAnimacion(FxData(B).Animacion, Val(tNBMP.text), DGH)
    Case 5
        If N > NumEscudosAnims Then
            NumEscudosAnims = NumEscudosAnims + 1
            ReDim Preserve MisEscudos(1 To NumEscudosAnims)
            ReDim Preserve ShieldAnimData(1 To NumEscudosAnims)
        End If
        For i = 1 To 4
            MisEscudos(N).Arma(i) = CopiarAnimacion(MisEscudos(B).Arma(i), Val(tNBMP.text), DGH)
            InitGrh ShieldAnimData(N).ShieldWalk(i), MisEscudos(N).Arma(i), 0
        Next i
End Select
If N > lstCual.ListCount Then
    lstCual.AddItem N
Else
    lstCual.List(N - 1) = N
End If
End Sub

Function CopiarAnimacion(GI As Integer, bmp As Integer, ByRef DGH As Integer) As Integer
Dim i As Integer
Dim Index As Integer
Dim X As Integer
Dim l As Integer
Dim xFrames() As Integer
    If GrhData(GI).NumFrames = 1 Then
        Index = AgregarGrh(i, False, DGH)
        If DGH > 0 Then DGH = DGH + 1
        'Fill in grh data
        With GrhData(Index)
            .FileNum = bmp
            
            ReDim .Frames(1 To 1) As Integer
            .Frames(1) = Index
            
            .NumFrames = 1
            .pixelHeight = GrhData(GI).pixelHeight
            .pixelWidth = GrhData(GI).pixelWidth
            .Speed = 0
            .sX = GrhData(GI).sX
            .sY = GrhData(GI).sY
            .TileHeight = GrhData(GI).TileHeight
            .TileWidth = GrhData(GI).TileWidth
        End With
        
        CopiarAnimacion = Index
    Else

            ReDim xFrames(1 To GrhData(GI).NumFrames) As Integer
            For X = 1 To GrhData(GI).NumFrames
                If DGH = 0 Then
                l = AgregarGrh(l, False, 0)
                Else
                
                l = AgregarGrh(l, False, DGH)
                DGH = DGH + 1
                End If
                    xFrames(X) = l
                    GrhData(l).FileNum = bmp
                    GrhData(l).NumFrames = 1
                    ReDim GrhData(l).Frames(1 To 1) As Integer
                    GrhData(l).Frames(1) = l
                    GrhData(l).pixelHeight = GrhData(GrhData(GI).Frames(X)).pixelHeight
                    GrhData(l).pixelWidth = GrhData(GrhData(GI).Frames(X)).pixelWidth
                    GrhData(l).Speed = 0
                    GrhData(l).sX = GrhData(GrhData(GI).Frames(X)).sX
                    GrhData(l).sY = GrhData(GrhData(GI).Frames(X)).sY
                    GrhData(l).TileHeight = GrhData(GrhData(GI).Frames(X)).TileHeight
                    GrhData(l).TileWidth = GrhData(GrhData(GI).Frames(X)).TileWidth
     
            Next X
            If DGH = 0 Then
                Index = AgregarGrh(i, True, 0)
            Else
                Index = AgregarGrh(i, True, DGH)
                DGH = DGH + 1
            End If
            GrhData(Index).FileNum = 0
            
            GrhData(Index).NumFrames = GrhData(GI).NumFrames
            ReDim GrhData(Index).Frames(1 To GrhData(Index).NumFrames) As Integer
            
            GrhData(Index).Frames = xFrames
            
            GrhData(Index).pixelHeight = GrhData(GI).pixelHeight
            GrhData(Index).pixelWidth = GrhData(GI).pixelWidth
            GrhData(Index).Speed = GrhData(GI).Speed
            GrhData(Index).sX = GrhData(GI).sX
            GrhData(Index).sY = GrhData(GI).sY
            GrhData(Index).TileHeight = GrhData(GI).TileHeight
            GrhData(Index).TileWidth = GrhData(GI).TileWidth
     

        CopiarAnimacion = Index
        
    End If
    
    
    

    
    'Now select it in the list
    'frmMain.grhList.ListIndex = i
End Function
Public Function AgregarGrh(ByRef EnLista As Integer, Optional Animacion As Boolean = False, Optional DGRH As Integer = 0) As Integer
    Dim Index As Integer
    Dim i As Integer
    
If DGRH = 0 Then
    Index = UBound(GrhData()) + 1
Else
    Index = DGRH
End If
    
    'Make sure he is not overwritting anything
    If Index <= UBound(GrhData()) Then
        If GrhData(Index).NumFrames > 0 Then
            'If MsgBox("The chosen index is currently in use. Do you want to overwrite it?", vbOKCancel) = vbCancel Then
            '    Exit Function
            'End If
        End If
    Else
        'Resize array
        ReDim Preserve GrhData(1 To Index) As GrhData
    End If
    
    If GrhData(Index).NumFrames = 0 Then
        'Search where to place the grh....
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) > Index Then
                Exit For
            End If
        Next i
        
        'Add it!
        Call frmMain.grhList.AddItem(Index & IIf(Animacion, " (ANIMACIÓN)", ""), i)
    Else
        'Search for the grh index within the grhList
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) = Index Then
                Exit For
            End If
        Next i
    End If
    
    AgregarGrh = Index
End Function

Private Sub Command19_Click()
Dim Index As Integer
Dim Cant As Integer
Dim i As Integer
Index = AgregarGrh(0, True, Val(tEnGrh.text))
Cant = Val(tAnimHasta.text) - Val(tAnimDesde.text) + 1
With GrhData(Index)
    .NumFrames = Cant
    ReDim .Frames(1 To Cant)
    For i = Val(tAnimDesde.text) To Val(tAnimHasta.text)
        .Frames(i - Val(tAnimDesde.text) + 1) = i
    Next i
    .Speed = Val(Text1.text) * Cant
End With
End Sub

Private Sub Command2_Click()
Dim grhindex As Integer

        If GrhData(currentGrh).NumFrames > 1 Then
            'Move to next animation frame!

            
            grhindex = GrhData(currentGrh).Frames(CurrentFrame)
        Else
            grhindex = currentGrh
        End If

GrhData(grhindex).sY = GrhData(grhindex).sY + 1
DrawFrame (CurrentFrame)
End Sub

Private Sub Command20_Click()
Dim N As Integer

Dim i As Integer
N = Val(tNext.text)
If N = 0 Then Exit Sub
Select Case Cual
    Case 0
        If N > NumBodies Then
            NumBodies = NumBodies + 1
            ReDim Preserve MisCuerpos(1 To NumBodies)
            ReDim Preserve BodyData(1 To NumBodies)
        End If
        MisCuerpos(N).HeadOffsetX = Val(tOffX.text)
        MisCuerpos(N).HeadOffsetY = Val(tOffY.text)
        BodyData(N).HeadOffset.X = Val(tOffX.text)
        BodyData(N).HeadOffset.Y = Val(tOffY.text)
        For i = 1 To 4
            MisCuerpos(N).Body(i) = Val(txtAnim(i - 1).text)
            InitGrh BodyData(N).Walk(i), MisCuerpos(N).Body(i), 0
        Next i

    Case 1
        If N > Numheads Then
            Numheads = Numheads + 1
            ReDim Preserve HeadData(1 To Numheads)
            ReDim Preserve MisCabezas(1 To Numheads)
        End If
        For i = 1 To 4
            MisCabezas(N).Head(i) = Val(txtAnim(i - 1).text)
            InitGrh HeadData(N).Head(i), MisCabezas(N).Head(i), 0
        Next i
    Case 2
        If N > NumCascos Then
            NumCascos = NumCascos + 1
            ReDim Preserve MisCascos(1 To NumCascos) As tIndiceCabeza
            ReDim Preserve CascoAnimData(1 To NumCascos) As HeadData
        End If
        For i = 1 To 4
            MisCascos(N).Head(i) = Val(txtAnim(i - 1).text)
            InitGrh CascoAnimData(N).Head(i), MisCascos(N).Head(i), 0
        Next i
    Case 3
        If N > NumWeaponAnims Then
            NumWeaponAnims = NumWeaponAnims + 1
            ReDim Preserve MisArmas(1 To NumWeaponAnims)
            ReDim Preserve WeaponAnimData(1 To NumWeaponAnims)
        End If
        For i = 1 To 4
            MisArmas(N).Arma(i) = Val(txtAnim(i - 1).text)
            InitGrh WeaponAnimData(N).WeaponWalk(i), MisArmas(N).Arma(i), 0
        Next i
    Case 4
        If N > NumFxs Then
            NumFxs = NumFxs + 1
            ReDim Preserve FxData(1 To NumFxs)
        End If
        FxData(N).OFFSETX = Val(tOffX.text)
        FxData(N).OFFSETY = Val(tOffY.text)
        FxData(N).Animacion = Val(txtAnim(0).text)
    Case 5
        If N > NumEscudosAnims Then
            NumEscudosAnims = NumEscudosAnims + 1
            ReDim Preserve MisEscudos(1 To NumEscudosAnims)
            ReDim Preserve ShieldAnimData(1 To NumEscudosAnims)
        End If
        For i = 1 To 4
            MisEscudos(N).Arma(i) = Val(txtAnim(i - 1).text)
            InitGrh ShieldAnimData(N).ShieldWalk(i), MisEscudos(N).Arma(i), 0
        Next i
End Select
If N > lstCual.ListCount Then
    lstCual.AddItem N
Else
    lstCual.List(N - 1) = N
End If
End Sub

Private Sub Command21_Click()
CargarDats
End Sub

Private Sub Command22_Click()

Dim W As Integer
Dim H As Integer
Dim X As Integer
Dim Y As Integer
Dim iList As Integer
Dim GI As Integer
Dim Indice As Integer
Dim DGRH As Integer
W = Val(tITAncho.text)
H = Val(tITLargo.text)

DGRH = Val(tITDesde.text)

For Y = 1 To H
    For X = 1 To W
        GI = AgregarGrh(iList, False, DGRH)
        If DGRH > 0 Then DGRH = DGRH + 1
        With GrhData(GI)
            .FileNum = Val(tITbmp.text)
            .sX = Val(tITX.text) * 32 * W + (X - 1) * 32
            .sY = Val(tITY.text) * 32 * H + (Y - 1) * 32
            .NumFrames = 1
            ReDim .Frames(1 To 1)
            .Frames(1) = GI
            .pixelWidth = 32
            .pixelHeight = 32
            .TileWidth = 32
            .TileHeight = 32
        End With
        
        If X = 1 And Y = 1 Then
            Indice = GI
        End If
    Next X
Next Y


MaxSup = MaxSup + 1
ReDim Preserve SupData(MaxSup) As SupData
With SupData(MaxSup)
    .Grh = Indice
    .Block = False
    .Capa = 0
    .Width = W
    .Height = H
    .name = tITNombre.text
End With
End Sub

Private Sub Command23_Click()
Dim i As Integer
Dim GI As Integer
Dim cG As Integer
Dim X As Integer
If MsgBox("¿Esta seguro que desea borrar la indexacion y el bmp de todos los elementos filtrados?", vbExclamation + vbYesNo) = vbYes Then

For X = 0 To lstBuscar.ListCount - 1
cG = Val(lstBuscar.List(X))
    If GrhData(currentGrh).NumFrames > 0 Then
        For i = 1 To GrhData(cG).NumFrames
            GI = GrhData(cG).Frames(i)
            BorrarGrh (GI)
        Next i
        BorrarGrh (cG)
    Else
        BorrarGrh (cG)
    End If
Next X
End If
End Sub

Private Sub Command24_Click()
'khalem
lstCual.Visible = False
grhList.Visible = True
Dim i As Integer
For i = 0 To 3
optDir(i).Enabled = False
Next i
End Sub

Private Sub Command25_Click()
Dim FSO As New FileSystemObject
Dim X As file
Dim Num As Integer
Dim F As Folder
Dim i As Integer
Dim G As Integer
Dim e As Integer
Dim Encontro As Boolean
Set F = FSO.GetFolder(Config.bmpPath)
PB.max = F.Files.Count
For Each X In F.Files
    Num = Val(X.name)
    PB.value = i
    i = i + 1
    If Num > 0 Then
        G = BMPenGRH(Num)
        If G = 0 Then
            FSO.CopyFile F.path & "\" & Num & ".bmp", App.path & "\Categorias\NoUsados\" & Num & ".bmp"
        Else
            Encontro = False
            For e = 0 To UBound(BodyData)
                If BodyData(e).Walk(1).grhindex = G Or BodyData(e).Walk(2).grhindex = G Or BodyData(e).Walk(3).grhindex = G Or BodyData(e).Walk(4).grhindex = G Then
                    FSO.CopyFile F.path & "\" & Num & ".bmp", App.path & "\Categorias\Cuerpos\" & Num & ".bmp"
                    Encontro = True
                    Exit For
                End If
            Next e
            For e = 0 To UBound(HeadData)
                If HeadData(e).Head(1).grhindex = G Or HeadData(e).Head(2).grhindex = G Or HeadData(e).Head(3).grhindex = G Or HeadData(e).Head(4).grhindex = G Then
                    FSO.CopyFile F.path & "\" & Num & ".bmp", App.path & "\Categorias\Cabezas\" & Num & ".bmp"
                    Encontro = True
                    Exit For
                End If
            Next e
            For e = 0 To UBound(CascoAnimData)
                If CascoAnimData(e).Head(1).grhindex = G Or CascoAnimData(e).Head(2).grhindex = G Or CascoAnimData(e).Head(3).grhindex = G Or CascoAnimData(e).Head(4).grhindex = G Then
                    FSO.CopyFile F.path & "\" & Num & ".bmp", App.path & "\Categorias\Cascos\" & Num & ".bmp"
                    Encontro = True
                    Exit For
                End If
            Next e
            For e = 0 To UBound(WeaponAnimData)
                If WeaponAnimData(e).WeaponWalk(1).grhindex = G Or WeaponAnimData(e).WeaponWalk(2).grhindex = G Or WeaponAnimData(e).WeaponWalk(3).grhindex = G Or WeaponAnimData(e).WeaponWalk(4).grhindex = G Then
                    FSO.CopyFile F.path & "\" & Num & ".bmp", App.path & "\Categorias\Armas\" & Num & ".bmp"
                    Encontro = True
                    Exit For
                End If
            Next e
            For e = 0 To UBound(ShieldAnimData)
                If ShieldAnimData(e).ShieldWalk(1).grhindex = G Or ShieldAnimData(e).ShieldWalk(2).grhindex = G Or ShieldAnimData(e).ShieldWalk(3).grhindex = G Or ShieldAnimData(e).ShieldWalk(4).grhindex = G Then
                    FSO.CopyFile F.path & "\" & Num & ".bmp", App.path & "\Categorias\Escudos\" & Num & ".bmp"
                    Encontro = True
                    Exit For
                End If
            Next e
            For e = 0 To UBound(FxData)
                If FxData(e).Animacion = G Then
                    FSO.CopyFile F.path & "\" & Num & ".bmp", App.path & "\Categorias\FX\" & Num & ".bmp"
                    Encontro = True
                    Exit For
                End If
            Next e
            For e = 0 To UBound(ObjData)
                If ObjData(e).grhindex = G Then
                    FSO.CopyFile F.path & "\" & Num & ".bmp", App.path & "\Categorias\Obj\" & Num & ".bmp"
                    Encontro = True
                    Exit For
                End If
            Next e
            
            If Not Encontro Then
                FSO.CopyFile F.path & "\" & Num & ".bmp", App.path & "\Categorias\Otros\" & Num & ".bmp"
            End If
        End If
    End If
Next X
End Sub

Private Sub Command26_Click()
MaxSup = MaxSup + 1
ReDim Preserve SupData(MaxSup) As SupData
With SupData(MaxSup)
    .Grh = Val(tITSolo.text)
    .Block = False
    .Capa = 0
    .Width = 1
    .Height = 1
    .name = tITNombre.text
End With
End Sub

Private Sub Command27_Click()
frmNPCS.Show
End Sub

Private Sub Command3_Click()
Dim grhindex As Integer

        If GrhData(currentGrh).NumFrames > 1 Then

            grhindex = GrhData(currentGrh).Frames(CurrentFrame)
        Else
            grhindex = currentGrh
        End If

GrhData(grhindex).sX = GrhData(grhindex).pixelHeight + 1
DrawFrame (CurrentFrame)
End Sub

Private Sub Command4_Click()
Dim grhindex As Integer

        If GrhData(currentGrh).NumFrames > 1 Then

            grhindex = GrhData(currentGrh).Frames(CurrentFrame)
        Else
            grhindex = currentGrh
        End If

GrhData(grhindex).pixelHeight = GrhData(grhindex).pixelHeight - 1
DrawFrame (CurrentFrame)
End Sub

Private Sub Command5_Click()
Dim grhindex As Integer

        If GrhData(currentGrh).NumFrames > 1 Then

            grhindex = GrhData(currentGrh).Frames(CurrentFrame)
        Else
            grhindex = currentGrh
        End If

GrhData(grhindex).pixelHeight = GrhData(grhindex).sX + 1
DrawFrame (CurrentFrame)
End Sub

Private Sub Command6_Click()
Dim grhindex As Integer

        If GrhData(currentGrh).NumFrames > 1 Then

            grhindex = GrhData(currentGrh).Frames(CurrentFrame)
        Else
            grhindex = currentGrh
        End If

GrhData(grhindex).sX = GrhData(grhindex).sX - 1
DrawFrame (CurrentFrame)
End Sub

Private Sub Command7_Click()
Dim grhindex As Integer

        If GrhData(currentGrh).NumFrames > 1 Then

            grhindex = GrhData(currentGrh).Frames(CurrentFrame)
        Else
            grhindex = currentGrh
        End If

GrhData(grhindex).pixelWidth = GrhData(grhindex).pixelWidth + 1
DrawFrame (CurrentFrame)
End Sub

Private Sub Command8_Click()
Dim grhindex As Integer

        If GrhData(currentGrh).NumFrames > 1 Then

            grhindex = GrhData(currentGrh).Frames(CurrentFrame)
        Else
            grhindex = currentGrh
        End If

GrhData(grhindex).pixelWidth = GrhData(grhindex).pixelWidth - 1
DrawFrame (CurrentFrame)
End Sub

Private Sub Command9_Click()
        If GrhData(currentGrh).NumFrames > 1 Then
            'Move to next animation frame!
            CurrentFrame = CurrentFrame - 1
            
            If CurrentFrame < 1 Then
                CurrentFrame = GrhData(currentGrh).NumFrames
            End If
            
            DrawFrame (CurrentFrame)
        End If
End Sub


Private Sub currentPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lPosMouse.Caption = "X: " & X & "   -   Y: " & Y
End Sub


Private Sub Commanda_MouseDown( _
        Index As Integer, _
        Button As Integer, _
        Shift As Integer, _
        X As Single, Y As Single)

    'Coloca el Flag en True
    Mouse_Abajo = True

    Select Case Index
        Case 0
            'Desplaza la imagen hacia la derecha
            While Pos_x > 0 And Mouse_Abajo
                Pos_x = Pos_x - Avance
                currentPic.PaintPicture currentPic.Picture, 0, 0, , , _
                                      Pos_x, Pos_y, currentPic.ScaleWidth, _
                                      currentPic.ScaleHeight
               
                DoEvents
            Wend

        Case 1
            'Desplaza la imagen hacia arriba
            While Pos_y < Alto_Pic And Mouse_Abajo

                Pos_y = Pos_y + Avance
                currentPic.PaintPicture currentPic.Picture, 0, 0, , , Pos_x, Pos_y, _
                                    currentPic.ScaleWidth, currentPic.ScaleHeight

                DoEvents
            Wend
            'Desplaza la imagen hacia la izquierda
        Case 2
            While Pos_x < Ancho_Pic And Mouse_Abajo
                
                Pos_x = Pos_x + Avance
                currentPic.PaintPicture currentPic.Picture, 0, 0, , , Pos_x, _
                                    Pos_y, ScaleWidth, currentPic.ScaleHeight
               
                DoEvents
            Wend
            'Desplaza la imagen hacia abajo
        Case 3
        
        While Pos_y > 0 And Mouse_Abajo
                
                Pos_y = Pos_y - Avance
                currentPic.PaintPicture currentPic.Picture, 0, 0, , , Pos_x, Pos_y, _
                       currentPic.ScaleWidth, currentPic.ScaleHeight
                
                DoEvents
        Wend

    End Select

End Sub

Private Sub Commanda_MouseUp(Index As Integer, Button As Integer, _
                        Shift As Integer, X As Single, Y As Single)

    ' cuando se suelta el mouse en el Command finalizar
    Mouse_Abajo = False
End Sub


Private Sub Form_Load()
Avance = 1

    Dim i As Long
    Dim fileName As String
    Dim path As String
    
    
    DirX = 1
    Me.Show
    If Not LoadConfig() Then
        'Show config form
        Call frmConfig.Show(vbModal, Me)
    End If
    
    'Load Grhs!
    Call LoadGrhData(Config.initPath)
    
    
    
    'Set up bmp search path
    If Right$(Config.bmpPath, 1) <> "\" Then
        path = Config.bmpPath & "\*.bmp"
    Else
        path = Config.bmpPath & "*.bmp"
    End If
    

    
    'Set default zoom value
    ZoomTxt.text = DEFAULT_ZOOM
    
    editionCoord = sbpeNone
    
    currentGrh = NO_GRH
    
    'By default update events are not ignored
    ignoreGrhTextUpdate = False
    
    'Show first grh by default
    If grhList.ListCount > 0 Then
        grhList.ListIndex = 0
    End If
    
    CargarOtrosInit
    DoEvents
    CargarLista
End Sub
Sub CargarLista(Optional Filtrado As Boolean = False)
    'Fill the lists
    Dim i As Integer
    'Exit Sub
    PB.max = UBound(GrhData())
    
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
        If Filtrado Then
        
            
            If Not EstaEnUso(i) Then
                If GrhData(i).NumFrames = 1 Then
                    Call grhList.AddItem(CStr(i))
                Else
                    Call grhList.AddItem(CStr(i) & " (ANIMACIÓN)")
                End If
            End If

        
         Else
                If GrhData(i).NumFrames = 1 Then
                    Call grhList.AddItem(CStr(i))
                Else
                    Call grhList.AddItem(CStr(i) & " (ANIMACIÓN)")
                End If
        End If
        End If
        PB.value = i
        If PB.value = PB.max Then
        PB.Visible = False
        End If
    Next i
End Sub
Private Sub grhHeightTxt_Change()
If FinCarga Then
    'Prevent non numeric characters
    If Not IsNumeric(grhHeightTxt.text) Then
        grhHeightTxt.text = Val(grhHeightTxt.text)
    End If
    
    'Prevent overflow
    If Val(grhHeightTxt.text) > &H7FFF Then
        grhHeightTxt.text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    'If CInt(grhHeightTxt.Text) > previewer.ScaleY(currentPic.Height) - Val(grhYTxt.Text) Then
    '    grhHeightTxt.Text = Round(previewer.ScaleY(currentPic.Height) - Val(grhYTxt.Text))
    'End If
    
    'Prevent negative values
    If CInt(grhHeightTxt.text) < 0 Then
        grhHeightTxt.text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).pixelHeight = CInt(grhHeightTxt.text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, CurrentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaEndY = selectionAreaStartY + Val(grhHeightTxt.text)
    
    'Redraw selection area
    Call RenderSelectionBox
    End If
End Sub

Private Sub grhList_Click()

    ' Set current grh and reset frame
    
    AbrirGrh (Val(grhList.text))
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
    
    'Should grh controls be enabled?
    Call SetGrhControlsEnabled(True)
    
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
    
    With currentPic
            'ancho y alto del gráfico cargado  en el Picture
        Ancho_Pic = CInt(.ScaleX(.Picture.Width, vbHimetric, .ScaleMode))
        Alto_Pic = CInt(.ScaleY(.Picture.Height, vbHimetric, .ScaleMode))
   End With
    End If
    'Enable animations if necessary
    If GrhData(currentGrh).NumFrames > 1 Then
        animation.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)
        lSpeed.Caption = "Speed: " & Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)
        animation.Enabled = True
        
        grhOnly.value = vbChecked
        grhOnly.Enabled = False
    Else
        animation.Enabled = False
        
        If Not grhOnly.Enabled Then
            grhOnly.Enabled = True
            
            'grhOnly.value = vbChecked
        ElseIf grhOnly.value = vbUnchecked Then
            'Set selection box!
            Call SelectGrhArea(currentGrh)
        End If
        
        'Show bmp
        bmpTxt.text = GrhData(currentGrh).FileNum
        
        'Filelist will reset the currentGrh, restore it!
        'currentGrh = Val(grhList.Text)
        
        'Set selection box!
        Call SelectGrhArea(currentGrh)
        
        'Display grh info
        grhXTxt.text = GrhData(currentGrh).sX
        grhYTxt.text = GrhData(currentGrh).sY
        grhWidthTxt.text = GrhData(currentGrh).pixelWidth
        grhHeightTxt.text = GrhData(currentGrh).pixelHeight
        
        selectionAreaStartX = GrhData(currentGrh).sX
        selectionAreaStartY = GrhData(currentGrh).sY
        selectionAreaEndX = GrhData(currentGrh).sX + GrhData(currentGrh).pixelWidth
        selectionAreaEndY = GrhData(currentGrh).sY + GrhData(currentGrh).pixelHeight
    End If
    
    'Set scrollers!
    'Call SetScrollers
    previewer.Width = GrhData(currentGrh).pixelWidth
    previewer.Height = GrhData(currentGrh).pixelHeight
    'Display the grh!
    Call RedrawPicture(currentGrh, CurrentFrame)

    'Show selection box (if needed)
    ignoreSelectionBoxRender = (grhOnly.value = vbChecked)
    
    FinCarga = True
    'Call RenderSelectionBox
End Sub
Private Sub grhOnly_Click()
    If currentGrh = NO_GRH Then Exit Sub
    
    Call RedrawPicture(currentGrh, CurrentFrame)
    
    ignoreSelectionBoxRender = (grhOnly.value = vbChecked)
    
    'Set selection box!
    Call SelectGrhArea(currentGrh)
    
    Call RenderSelectionBox
End Sub

Private Sub grhWidthTxt_Change()
If Not FinCarga Then Exit Sub
    'Prevent non numeric characters
    If Not IsNumeric(grhWidthTxt.text) Then
        grhWidthTxt.text = Val(grhWidthTxt.text)
    End If
    
    'Prevent overflow
    If Val(grhWidthTxt.text) > &H7FFF Then
        grhWidthTxt.text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    'If CInt(grhWidthTxt.Text) > previewer.ScaleX(currentPic.Width) - Val(grhXTxt.Text) Then
    '    grhWidthTxt.Text = Round(previewer.ScaleX(currentPic.Width) - Val(grhXTxt.Text))
    'End If
    
    'Prevent negative values
    If CInt(grhWidthTxt.text) < 0 Then
        grhWidthTxt.text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).pixelWidth = CInt(grhWidthTxt.text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, CurrentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaEndX = selectionAreaStartX + CInt(grhWidthTxt.text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub grhXTxt_Change()
If Not FinCarga Then Exit Sub
    'Prevent non numeric characters
    If Not IsNumeric(grhXTxt.text) Then
        grhXTxt.text = Val(grhXTxt.text)
    End If
    
    'Prevent overflow
    If Val(grhXTxt.text) > &H7FFF Then
        grhXTxt.text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    'If CInt(grhXTxt.Text) > previewer.ScaleX(currentPic.Width) Then
    '    grhXTxt.Text = Round(previewer.ScaleX(currentPic.Width))
    'End If
    
    'Prevent negative values
    If CInt(grhXTxt.text) < 0 Then
        grhXTxt.text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).sX = CInt(grhXTxt.text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, CurrentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaStartX = CInt(grhXTxt.text)
    selectionAreaEndX = selectionAreaStartX + Val(grhWidthTxt.text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub grhXTxt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then mAddgrh_Click
End Sub

Private Sub grhYTxt_Change()
If Not FinCarga Then Exit Sub
    'Prevent non numeric characters
    If Not IsNumeric(grhYTxt.text) Then
        grhYTxt.text = Val(grhYTxt.text)
    End If
    
    'Prevent overflow
    If Val(grhYTxt.text) > &H7FFF Then
        grhYTxt.text = &H7FFF
    End If
    
    'Prevent values way too big for the current bmp
    'If CInt(grhYTxt.Text) > previewer.ScaleY(currentPic.Height) Then
   '     grhYTxt.Text = Round(previewer.ScaleY(currentPic.Height))
    'End If
    
    'Trim height to prevent invalid values
    'If CInt(grhYTxt.Text) + Val(grhHeightTxt.Text) > previewer.ScaleY(currentPic.Height) Then
    '    grhHeightTxt.Text = Round(previewer.ScaleY(currentPic.Height)) - CInt(grhYTxt.Text)
    'End If
    
    'Prevent negative values
    If CInt(grhYTxt.text) < 0 Then
        grhYTxt.text = 0
    End If
    
    'Update data in memory
    If currentGrh <> NO_GRH Then
        GrhData(currentGrh).sY = CInt(grhYTxt.text)
        
        'Re-render updated grh
        Call RedrawPicture(currentGrh, CurrentFrame)
    End If
    
    'If an ignore was set, we end here
    If ignoreGrhTextUpdate Then Exit Sub
    
    'Set the selection are coord appropiately
    selectionAreaStartY = Val(grhYTxt.text)
    selectionAreaEndY = selectionAreaStartY + Val(grhHeightTxt.text)
    
    'Redraw selection area
    Call RenderSelectionBox
End Sub

Private Sub grhYTxt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then mAddgrh_Click
End Sub

Private Sub lstBuscar_Click()
AbrirGrh (Val(lstBuscar.text))
End Sub

Private Sub lstCual_Click()
On Local Error Resume Next
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

Private Sub mAddgrh_Click()
    Dim Index As Long
    Dim i As Long

    Index = UBound(GrhData()) + 1

    ReDim Preserve GrhData(1 To Index) As GrhData

    
    If GrhData(Index).NumFrames = 0 Then
        'Search where to place the grh....
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) > Index Then
                Exit For
            End If
        Next i
        
        'Add it!
        Call frmMain.grhList.AddItem(Index, i)
    Else
        'Search for the grh index within the grhList
        For i = 0 To frmMain.grhList.ListCount - 1
            If Val(frmMain.grhList.List(i)) = Index Then
                Exit For
            End If
        Next i
    End If
    
    'Fill in grh data
    With GrhData(Index)
        .FileNum = Val(frmMain.bmpTxt.text)
        
        ReDim .Frames(1 To 1) As Integer
        .Frames(1) = Index
        
        .NumFrames = 1
        .pixelHeight = Val(frmMain.grhHeightTxt.text)
        .pixelWidth = Val(frmMain.grhWidthTxt.text)
        .Speed = 0
        .sX = Val(frmMain.grhXTxt.text)
        .sY = Val(frmMain.grhYTxt.text)
        .TileHeight = .pixelHeight / Config.TilePixelHeight
        .TileWidth = .pixelWidth / Config.TilePixelWidth
    End With
    
    'Now select it in the list
    frmMain.grhList.ListIndex = i
End Sub

Private Sub mCopiar_Click()
frmCopiar.Show vbModal
End Sub

Private Sub mCrearA_Click()
frmCrearAnim.Show
End Sub

Private Sub MMCB_Click()
Dim Datos As String
Datos = Clipboard.GetText

Dim Lineas() As String
Dim Index As Long

Lineas = Split(Datos, vbCrLf)

Dim tmpStr As String
Dim i As Integer

For i = 0 To UBound(Lineas)
    'ReadField
    tmpStr = ReadField(1, Lineas(i), Asc("="))
    Index = 17114 + Right$(tmpStr, Len(tmpStr) - 3)
    ReDim Preserve GrhData(1 To Index)
    With GrhData(Index)
        .FileNum = Val(ReadField(2, Lineas(i), 45)) + 14285
        
        ReDim .Frames(1 To 1) As Integer
        .Frames(1) = Index
        
        .NumFrames = 1
        .pixelHeight = Val(ReadField(6, Lineas(i), 45))
        .pixelWidth = Val(ReadField(5, Lineas(i), 45))
        .Speed = 0
        .sX = Val(ReadField(3, Lineas(i), 45))
        .sY = Val(ReadField(4, Lineas(i), 45))
        .TileHeight = .pixelHeight / Config.TilePixelHeight
        .TileWidth = .pixelWidth / Config.TilePixelWidth
    End With
    Call frmMain.grhList.AddItem(Index, i)
Next i



End Sub

Private Sub optDir_Click(Index As Integer)
DirX = Index + 1
lstCual_Click
End Sub

Private Sub picScrollH_Change()
    'Redraw
    Call RedrawPicture(currentGrh, CurrentFrame)
    
    'Show selection box!
    Call RenderSelectionBox
End Sub

Private Sub picScrollV_Change()
    'Redraw
    Call RedrawPicture(currentGrh, CurrentFrame)
    
    'Show selection box!
    Call RenderSelectionBox
End Sub

Private Sub previewer_Click()

Call SavePicture(previewer.Image, App.path & "\" & currentGrh & ".bmp")
End Sub

Private Sub previewer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If no picture is loaded, there is nothing to be done
    If currentPic.Picture Is Nothing Then Exit Sub
    
    If Button And vbLeftButton Then
        If currentGrh <> NO_GRH And grhOnly.value = vbChecked Then Exit Sub
        

    End If
End Sub

Private Sub RemoveGrhMnu_Click()
    Dim i As Long
    
    If currentGrh = NO_GRH Then
        MsgBox "There is no grh selected."
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete the grh " & currentGrh & "?" & vbCrLf & "This change can't be undone.", vbOKCancel) = vbOK Then
        'Reset it
        With GrhData(currentGrh)
            .FileNum = 0
            ReDim .Frames(0)
            .NumFrames = 0
            .pixelHeight = 0
            .pixelWidth = 0
            .Speed = 0
            .sX = 0
            .sY = 0
            .TileHeight = 0
            .TileWidth = 0
        End With
        
        'Remove it
        For i = 0 To grhList.ListCount - 1
            If Val(grhList.List(i)) = currentGrh Then
                grhList.RemoveItem (i)
                Exit For
            End If
        Next i
        
        'Select next grh
        If i < grhList.ListCount Then
            grhList.ListIndex = i
        Else
            grhList.ListIndex = grhList.ListCount - 1
        End If
    End If
End Sub

Private Sub SaveMnu_Click()
    'Detect the original file format and save it
    If Grh.fileVersion = -1 Then
        If Not Grh.SaveGrhDataOld(Config.initPath) Then
            Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk, or you are using grh indexes above 32767, which are only supported in the new file format.")
        Else
            Call MsgBox("File succesfully written.")
        End If
    Else
        If Not Grh.SaveGrhDataNew(Config.initPath) Then
            Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
        Else
            Call MsgBox("File succesfully written.")
        End If
    End If
    
    GuardarOtrosInit
End Sub

Private Sub SaveNewMnu_Click()
    If Not Grh.SaveGrhDataNew(Config.initPath) Then
        Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk.")
    Else
        Call MsgBox("File succesfully written.")
    End If
End Sub

Private Sub SaveOldMnu_Click()
    If MsgBox("The old file format speed system is FPS based, animation's speed may be altered. Do you want to proceed?", vbYesNo) = vbYes Then
        If Not Grh.SaveGrhDataOld(Config.initPath) Then
            Call MsgBox("The file could not be saved. This could be caused due to lack of space on disk, or you are using grh indexes above 32767, which are only supported in the new file format.")
        Else
            Call MsgBox("File succesfully written.")
        End If
    End If
End Sub

Private Sub tITAncho_Change()
DibujarIt
End Sub

Private Sub tITbmp_Change()
On Error Resume Next
Dim bmp As Integer
bmp = Val(tITbmp.text)
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
Sub DibujarIt()
Dim X As Integer
Dim Y As Integer
X = Val(tITX.text) * 32 * Val(tITAncho.text)
Y = Val(tITY.text) * 32 * Val(tITLargo.text)
currentPic.Cls
currentPic.Line (X, Y)-(X + Val(tITAncho.text) * 32, Y + Val(tITLargo.text) * 32), , B
End Sub
Private Sub tITLargo_Change()
DibujarIt
End Sub

Private Sub tITX_Change()
DibujarIt
End Sub

Private Sub tITY_Change()
DibujarIt
End Sub

Private Sub tOffX_Change()
tOffY_Change
End Sub

Private Sub tOffY_Change()
Select Case Cual
    Case 0
        BodyData(lstCual.ListIndex + 1).HeadOffset.X = Val(tOffX.text)
        BodyData(lstCual.ListIndex + 1).HeadOffset.Y = Val(tOffY.text)
        MisCuerpos(lstCual.ListIndex + 1).HeadOffsetX = Val(tOffX.text)
        MisCuerpos(lstCual.ListIndex + 1).HeadOffsetY = Val(tOffY.text)
    Case 4
        FxData(lstCual.ListIndex + 1).OFFSETX = Val(tOffX.text)
        FxData(lstCual.ListIndex + 1).OFFSETY = Val(tOffY.text)
End Select
End Sub

Private Sub ZoomIn_Click()
    ZoomTxt.text = Val(ZoomTxt.text) + ZOOM_STEP
End Sub

Private Sub ZoomOut_Click()
    ZoomTxt.text = Val(ZoomTxt.text) - ZOOM_STEP
End Sub

Private Sub ZoomTxt_Change()
    'Validate
    If Not FinCarga Then Exit Sub
    If Not IsNumeric(ZoomTxt.text) Then
        ZoomTxt.text = DEFAULT_ZOOM
        Exit Sub
    End If
    
    If Val(ZoomTxt.text) > MAX_ZOOM Then
        ZoomTxt.text = MAX_ZOOM
        Exit Sub
    End If
    
    If Val(ZoomTxt.text) < MIN_ZOOM Then
        ZoomTxt.text = MIN_ZOOM
        Exit Sub
    End If
    
    'Recompute zoom
    zoom = CSng(ZoomTxt.text) / DEFAULT_ZOOM
    
    
    'Reset scrollbars
    Call SetScrollers
    
    'Redraw
    Call RedrawPicture(currentGrh, CurrentFrame)
    
    'Show selection box!
    Call RenderSelectionBox
End Sub

''
' Sets the scrollers' properties appropiately for the current picture loaded, zoom and value.

Private Sub SetScrollers()
    Dim oldMax As Integer
    



End Sub

''
' Renders the last laoded picture.
'
' @param    grh     The grh to be rendered within the loaded picture. Can be @code NO_GRH
' @param    frame   The frame of the grh to be rendered. Only important if grh is not @code NO_GRH

Private Sub RedrawPicture(ByVal Grh As Long, ByVal frame As Long)
    If currentPic.Picture Is Nothing Then Exit Sub
    On Error Resume Next
    'Clear picturebox
       'previewer.Picture = LoadPicture("")
    
    'Render!
    If Grh <> NO_GRH And grhOnly.value = vbChecked Then
        'Transform grh to actual frame grh.
        Grh = GrhData(Grh).Frames(frame)
        previewer.Cls
        previewer.Width = GrhData(Grh).pixelWidth
        previewer.Height = GrhData(Grh).pixelHeight
        Call TransparentBlt(previewer.hDC, 0, 0, GrhData(Grh).pixelWidth, GrhData(Grh).pixelHeight, currentPic.hDC, GrhData(Grh).sX, GrhData(Grh).sY, GrhData(Grh).pixelWidth, GrhData(Grh).pixelHeight, vbMagenta)
        currentPic.Cls
        currentPic.Line (GrhData(Grh).sX, GrhData(Grh).sY)-(GrhData(Grh).sX + GrhData(Grh).pixelWidth, GrhData(Grh).sY + GrhData(Grh).pixelHeight), , B
        'Call previewer.PaintPicture(currentPic, -picScrollH.value * zoom, -picScrollV.value * zoom, _
        '                            GrhData(Grh).pixelWidth * zoom, _
        '                            GrhData(Grh).pixelHeight * zoom, _
        '                            GrhData(Grh).sX, GrhData(Grh).sY, _
        '                            GrhData(Grh).pixelWidth, GrhData(Grh).pixelHeight)
    Else
    End If
End Sub

''
' Renders the selection box.

Private Sub RenderSelectionBox()
 
End Sub



''
' Updates the appropiate selection box coords according to the current value of @code editionCoord.
'
' @param    x   The mouse pos in the x coord within the previewer.
' @param    y   The mouse pos in the y coord within the previewer.

Private Sub UpdateSelectionBox(ByVal X As Long, ByVal Y As Long)
    Dim tmp As Long
    
    
    'Invert coordinates if needed to prevent pointer from going crazy on corners.
    If selectionAreaStartX > selectionAreaEndX Then
        tmp = selectionAreaEndX
        selectionAreaEndX = selectionAreaStartX
        selectionAreaStartX = tmp
        
        'Invert edition coord accordingly.
        Select Case editionCoord
            Case sbpeEndX
                editionCoord = sbpeStartX
            
            Case sbpeEndXEndY
                editionCoord = sbpeStartXEndY
            
            Case sbpeEndXStartY
                editionCoord = sbpeStartXStartY
            
            Case sbpeStartX
                editionCoord = sbpeEndX
            
            Case sbpeStartXEndY
                editionCoord = sbpeEndXEndY
            
            Case sbpeStartXStartY
                editionCoord = sbpeEndXStartY
        End Select
    End If
    
    If selectionAreaStartY > selectionAreaEndY Then
        tmp = selectionAreaEndY
        selectionAreaEndY = selectionAreaStartY
        selectionAreaStartY = tmp
        
        'Invert edition coord accordingly.
        Select Case editionCoord
            Case sbpeEndY
                editionCoord = sbpeStartY
            
            Case sbpeEndXEndY
                editionCoord = sbpeEndXStartY
            
            Case sbpeEndXStartY
                editionCoord = sbpeEndXEndY
            
            Case sbpeStartY
                editionCoord = sbpeEndY
            
            Case sbpeStartXEndY
                editionCoord = sbpeStartXStartY
            
            Case sbpeStartXStartY
                editionCoord = sbpeStartXEndY
        End Select
    End If
    
    'Display data at the bottom
    ignoreGrhTextUpdate = True
    
    grhHeightTxt.text = selectionAreaEndY - selectionAreaStartY
    grhWidthTxt.text = selectionAreaEndX - selectionAreaStartX
    grhXTxt.text = selectionAreaStartX
    grhYTxt.text = selectionAreaStartY
    
    ignoreGrhTextUpdate = False
End Sub

''
' Sets up the selection area around the given grh within it's bmp.
'
' @param    grh     The grh to be selected.

Private Sub SelectGrhArea(ByVal Grh As Long)
    selectionAreaStartX = GrhData(Grh).sX
    selectionAreaStartY = GrhData(Grh).sY
    selectionAreaEndX = selectionAreaStartX + GrhData(Grh).pixelWidth
    selectionAreaEndY = selectionAreaStartY + GrhData(Grh).pixelHeight
End Sub

''
'Enables / disables the grh controls (those within the grhFrame control).
'
' @param    enable  True if controls should be enabled, False otherwise.

Private Sub SetGrhControlsEnabled(ByVal enable As Boolean)
    Dim i As Long
    
    For i = 0 To frmMain.Controls.Count - 1
        If Not TypeOf frmMain.Controls(i) Is Timer And Not TypeOf frmMain.Controls(i) Is Menu Then
            If frmMain.Controls(i).Container Is grhFrame Then
                frmMain.Controls(i).Enabled = enable
            End If
        End If
    Next i
    
    grhFrame.Enabled = enable
End Sub
