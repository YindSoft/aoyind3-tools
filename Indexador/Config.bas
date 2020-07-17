Attribute VB_Name = "Config"
Option Explicit

Private Const CONFIG_FILE As String = "/Config.ini"

'Configuration variables are publicly accessed
Public bmpPath As String
Public initPath As String
Public datPath As String
Public indicePath As String

Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

''
' Loads the config file.
'
' @return   True upon succes, False otherwise.

Public Function LoadConfig() As Boolean
    
    Dim configPath As String
    
    configPath = App.path & CONFIG_FILE
    
    'Make sure the file exists
    If Not FileExists(configPath) Then
        Exit Function
    End If
    
    bmpPath = GetVar(configPath, "PATHS", "BMPPATH")
    initPath = GetVar(configPath, "PATHS", "INITPATH")
    datPath = GetVar(configPath, "PATHS", "DATPATH")
    indicePath = GetVar(configPath, "PATHS", "INDICEPATH")
    If Right$(initPath, 1) <> "\" Then initPath = initPath & "\"
    If Right$(datPath, 1) <> "\" Then datPath = datPath & "\"
    TilePixelHeight = Val(GetVar(configPath, "DEFINES", "TILEHEIGHT"))
    TilePixelWidth = Val(GetVar(configPath, "DEFINES", "TILEWIDTH"))
    
    'Make usre they are valid
    If bmpPath = "" Or Not DirExists(bmpPath) Or initPath = "" Or Not DirExists(initPath) _
            Or TilePixelHeight = 0 Or TilePixelWidth = 0 Then
        Exit Function
    End If
    
    LoadConfig = True
End Function

Public Sub SaveConfig()
    
    Dim configPath As String
    
    configPath = App.path & CONFIG_FILE
    
    Call WriteVar(configPath, "PATHS", "BMPPATH", bmpPath)
    Call WriteVar(configPath, "PATHS", "INITPATH", initPath)
    
    Call WriteVar(configPath, "DEFINES", "TILEHEIGHT", CStr(TilePixelHeight))
    Call WriteVar(configPath, "DEFINES", "TILEWIDTH", CStr(TilePixelWidth))
End Sub
