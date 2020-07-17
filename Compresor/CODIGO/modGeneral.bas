Attribute VB_Name = "modGeneral"
Option Explicit

Public Const GRAPHIC_PATH As String = "\GRAFICOS\"
Public Const RESOURCE_PATH As String = "\RECURSOS\"
Public Const PATCH_PATH As String = "\PARCHES\"
Public Const EXTRACT_PATH As String = "\EXTRACCIONES\"

'Public Declare Function GetTickCount Lib "kernel32" () As Long

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function
