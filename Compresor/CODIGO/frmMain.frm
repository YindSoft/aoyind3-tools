VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Compresor de recursos graficos"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton oInits 
      Caption         =   "Inits"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1695
   End
   Begin VB.OptionButton oWavs 
      Caption         =   "Wavs"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.OptionButton oInterfaces 
      Caption         =   "Interfaces"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.OptionButton oGraficos 
      Caption         =   "Graficos"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.Frame StatusFrame 
      Caption         =   "StatusFrame"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   5415
      Begin MSComctlLib.ProgressBar StatusBar 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton cmdPatch 
      Caption         =   "Parchear"
      Height          =   735
      Left            =   3720
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extraer"
      Height          =   735
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Text            =   "0"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Comprimir"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Working Version :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCompress_Click()
    Dim SourcePath As String
    Dim OutputPath As String
    
    If oGraficos.Value = True Then
        SourcePath = App.Path & GRAPHIC_PATH
        OutputPath = App.Path & RESOURCE_PATH & "Graficos\" & txtVersion.Text & "\"
        GRH_SOURCE_FILE_EXT = ".bmp"
        GRH_RESOURCE_FILE = "Graphics.AO"
        GRH_PATCH_FILE = "Graphics.PATCH"
    ElseIf oInterfaces.Value Then
        SourcePath = App.Path & "\Interface\"
        OutputPath = App.Path & RESOURCE_PATH & "Interface\" & txtVersion.Text & "\"
        GRH_SOURCE_FILE_EXT = ".*"
        GRH_RESOURCE_FILE = "Interface.AO"
        GRH_PATCH_FILE = "Interface.PATCH"
    ElseIf oWavs.Value Then
        SourcePath = App.Path & "\Wav\"
        OutputPath = App.Path & RESOURCE_PATH & "Wav\" & txtVersion.Text & "\"
        GRH_SOURCE_FILE_EXT = ".*"
        GRH_RESOURCE_FILE = "Wav.AO"
        GRH_PATCH_FILE = "Wav.PATCH"
    ElseIf oInits.Value Then
        SourcePath = App.Path & "\Inits\"
        OutputPath = App.Path & RESOURCE_PATH & "Init\" & txtVersion.Text & "\"
    End If
    
    
    
    
    
    'Check if the version already exists
    If FileExist(OutputPath & GRH_RESOURCE_FILE, vbNormal) Then
        If MsgBox("La versión ya se encuentra comprimida. ¿Desea reemplazarla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        If Not FileExist(OutputPath, vbDirectory) Then
            'Create this version folder
            MkDir OutputPath
        End If
    End If
    

    StatusFrame.Caption = "Comprimiendo..."
    
    'Compress!
    If Compress_Files(SourcePath, OutputPath, txtVersion.Text, StatusBar) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    

End Sub

Private Sub cmdExtract_Click()
    Dim ResourcePath As String
    Dim SourcePath As String
    Dim OutputPath As String
    
    
        If oGraficos.Value = True Then
        SourcePath = App.Path & GRAPHIC_PATH
        OutputPath = App.Path & EXTRACT_PATH & "Graficos\" & txtVersion.Text & "\"
        GRH_SOURCE_FILE_EXT = ".bmp"
        GRH_RESOURCE_FILE = "Graphics.AO"
        GRH_PATCH_FILE = "Graphics.PATCH"
    ElseIf oInterfaces.Value Then
        SourcePath = App.Path & "\Interface\"
        OutputPath = App.Path & EXTRACT_PATH & "Interface\" & txtVersion.Text & "\"
        GRH_SOURCE_FILE_EXT = ".*"
        GRH_RESOURCE_FILE = "Interface.AO"
        GRH_PATCH_FILE = "Interface.PATCH"
    ElseIf oWavs.Value Then
        SourcePath = App.Path & "\Wav\"
        OutputPath = App.Path & EXTRACT_PATH & "Wav\" & txtVersion.Text & "\"
        GRH_SOURCE_FILE_EXT = ".*"
        GRH_RESOURCE_FILE = "Wav.AO"
        GRH_PATCH_FILE = "Wav.PATCH"
    ElseIf oInits.Value Then
        SourcePath = App.Path & "\Inits\"
        OutputPath = App.Path & EXTRACT_PATH & "Init\" & txtVersion.Text & "\"
    End If

   
    'Check if the resource file exists
    If Not FileExist(ResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos a extraer." & vbCrLf & ResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("La versión ya se encuentra extraida. ¿Desea reextraerla?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        'Create this version folder
        MkDir OutputPath
    End If
    

    StatusFrame.Caption = "Extrayendo..."
    
    'Extract!
    If Extract_Files(ResourcePath, OutputPath, StatusBar) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    

End Sub

Private Sub cmdPatch_Click()
    Dim NewResourcePath As String
    Dim OldResourcePath As String
    Dim OutputPath As String
    Dim SourcePath As String
    Dim NewVersion As Long
    Dim OldVersion As Long
    
    NewVersion = CLng(txtVersion.Text)
    OldVersion = NewVersion - 1 'we patch from the last version
    
    If oGraficos.Value = True Then
        NewResourcePath = App.Path & RESOURCE_PATH & GRAPHIC_PATH & NewVersion & "\"
        OldResourcePath = App.Path & RESOURCE_PATH & GRAPHIC_PATH & OldVersion & "\"
        
        OutputPath = App.Path & EXTRACT_PATH & "Graficos\" & OldVersion & " to " & NewVersion & "\"
        GRH_SOURCE_FILE_EXT = ".bmp"
        GRH_RESOURCE_FILE = "Graphics.AO"
        GRH_PATCH_FILE = "Graphics.PATCH"
    ElseIf oInterfaces.Value Then
        NewResourcePath = App.Path & RESOURCE_PATH & "\Interface\" & NewVersion & "\"
        OldResourcePath = App.Path & RESOURCE_PATH & "\Interface\" & OldVersion & "\"
        OutputPath = App.Path & EXTRACT_PATH & "Interface\" & OldVersion & " to " & NewVersion & "\"
        GRH_SOURCE_FILE_EXT = ".*"
        GRH_RESOURCE_FILE = "Interface.AO"
        GRH_PATCH_FILE = "Interface.PATCH"
    ElseIf oWavs.Value Then
        NewResourcePath = App.Path & RESOURCE_PATH & "\Wav\" & NewVersion & "\"
        OldResourcePath = App.Path & RESOURCE_PATH & "\Wav\" & OldVersion & "\"
        OutputPath = App.Path & EXTRACT_PATH & "Wav\" & OldVersion & " to " & NewVersion & "\"
        GRH_SOURCE_FILE_EXT = ".*"
        GRH_RESOURCE_FILE = "Wav.AO"
        GRH_PATCH_FILE = "Wav.PATCH"
    ElseIf oInits.Value Then
        SourcePath = App.Path & "\Inits\"
        OutputPath = App.Path & PATCH_PATH & "Init\" & OldVersion & " to " & NewVersion & "\"
    End If
    
    
    
    NewResourcePath = App.Path & RESOURCE_PATH & NewVersion & "\"
    OldResourcePath = App.Path & RESOURCE_PATH & OldVersion & "\"
    OutputPath = App.Path & PATCH_PATH & OldVersion & " to " & NewVersion & "\"
    
    'Check if the new resource file exists
    If Not FileExist(NewResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos de la version actual." & vbCrLf & NewResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the old resource file exists
    If Not FileExist(OldResourcePath & GRH_RESOURCE_FILE, vbNormal) Then
        MsgBox "No se encontraron los recursos de la version anterior." & vbCrLf & OldResourcePath, , "Error"
        Exit Sub
    End If
    
    'Check if the version is already extracted
    If FileExist(OutputPath, vbDirectory) Then
        If MsgBox("El parche ya se ecnuentra realizado. ¿Desea reparchear?", vbYesNo, "Atencion") = vbNo Then _
            Exit Sub
    Else
        'Create this version folder
        MkDir OutputPath
    End If
    

    StatusFrame.Caption = "Armando el parche de " & OldVersion & " a " & NewVersion
    
    'Patch!
    If Make_Patch(NewResourcePath, OldResourcePath, OutputPath, StatusBar) Then
        'Show we finished
        MsgBox "Operación terminada con éxito"
    Else
        'Show we finished
        MsgBox "Operación abortada"
    End If
    

End Sub

