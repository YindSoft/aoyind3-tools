Attribute VB_Name = "Grh"
Option Explicit

Private Const GRH_DAT_FILE As String = "Graficos.ind"
Private Const OLD_FORMAT_HEADER As String = "Argentum Online by Noland-Studios."
Private Const OLD_FORMAT_INIT_FILE As String = "Inicio.con"

Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Integer
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Integer
    
    Speed As Single
End Type

Private Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    fX As Byte
    tip As Byte
    Password As String
    name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public GrhData() As GrhData

Public fileVersion As Long
Public grhCount As Long

Public Function LoadGrhData(ByVal path As String) As Boolean
On Error GoTo ErrHandler
    Dim handle As Integer
    Dim MiCabecera As tCabecera
    
    'Set initial size
    ReDim GrhData(0) As GrhData
    
    handle = FreeFile()
    
    If path = vbNullString Then Exit Function
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    If Not FileExists(path & GRH_DAT_FILE) Then
        MsgBox "The file " & path & GRH_DAT_FILE & " does not exist. A new one will be created with your work."
        Exit Function
    End If
    
    Open path & GRH_DAT_FILE For Binary Access Read Lock Write As handle
    
    'Check file format! (The crappy header had to have some use after all!)
    Get handle, , MiCabecera
    
    If Left$(MiCabecera.desc, Len(OLD_FORMAT_HEADER)) = OLD_FORMAT_HEADER Then
        LoadGrhData = LoadGrhDataOld(handle, NumberOfGrhs(path))
        
        'No version available in old file format
        fileVersion = -1
    Else
        'We dont' have header, move back to the beginning
        Seek handle, 1
        
        LoadGrhData = LoadGrhDataNew(handle)
    End If
    
    Close handle
Exit Function

ErrHandler:
    Close handle
    
    MsgBox "An error occured while loading the grh data." & vbCrLf _
        & "Make sure file format is valid, and in case of using the old format, make sure the " _
        & OLD_FORMAT_INIT_FILE & " file is in the init folder"
End Function
Function ReadField(ByVal Pos As Integer, ByRef text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(text, lastPos + 1, Len(text) - lastPos)
    Else
        ReadField = mid$(text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function
''
' Old crappy format loading. Restricted to 2^15-1 grhs,
' stores animation speed in frames and other crappy stuff.
' Coded just for backwards compatibility, users should avoid using this format.
'
' @param    handle      Handle to the open file containing the grh data.
'                       The header should have allready been removed.
' @param    totalGrhs   The total number of grhs that could exist.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhDataOld(ByVal handle As Integer, ByVal totalGrhs As Long) As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Integer
    Dim frame As Long
    Dim tempint As Integer
    Dim max As Integer
    
    max = -1
    
    'Resize array
    ReDim GrhData(1 To totalGrhs) As GrhData
    
    'Open files
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Fill Grh List
    
    'Get first Grh Number
    Get handle, , Grh
    
    Do Until Grh <= 0
        'Get highest grh number being used
        If Grh > max Then
            max = Grh
        End If
        
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            'Resize animation array
            ReDim .Frames(1 To .NumFrames) As Integer
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For frame = 1 To .NumFrames
                
                    Get handle, , tempint
                    
                    'Old format uses integers
                    .Frames(frame) = tempint
                    
                    If .Frames(frame) <= 0 Or .Frames(frame) > totalGrhs Then
                        GoTo ErrorHandler
                    End If
                Next frame
                
                Get handle, , tempint
                
                'Convert old speed to new one (time based)!
                .Speed = CSng(tempint) * .NumFrames * 1000 / 18
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , tempint
                
                'Old format used ints, not longs.
                .FileNum = tempint
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , .sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                    
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
        
        'Get Next Grh Number
        Get handle, , Grh
    Loop
    
    Close handle
    
    'Trim array
    ReDim Preserve GrhData(1 To max) As GrhData
    
    LoadGrhDataOld = True
Exit Function

ErrorHandler:
    LoadGrhDataOld = False
End Function

''
' Finds out the number of grhs for the old file format
'
' @param    path    The path to the folder in which the init file is stored.
'
' @return   The number of grhs that can exist at most.

Private Function NumberOfGrhs(ByVal path As String) As Long
    Dim N As Integer
    Dim GameIni As tGameIni
    Dim MiCabecera As tCabecera
    
    N = FreeFile
    
    Open path & OLD_FORMAT_INIT_FILE For Binary As #N
    
    Get N, , MiCabecera
    
    Get N, , GameIni
    
    Close N
    
    NumberOfGrhs = GameIni.NumeroDeBMPs
End Function

''
' Loads grh data using the new file format.
'
' @param    handle      Handle to the open file containing the grh data.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhDataNew(ByVal handle As Integer) As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim frame As Long
    Dim tmpLng As Long
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        If Grh > 0 Then
        
        If Grh = 16120 Then
        Beep
        End If
        
        
        With GrhData(Grh)
            'Get number of frames
            
            Get handle, , .NumFrames
            
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For frame = 1 To .NumFrames
                    Get handle, , .Frames(frame)
                    If .Frames(frame) <= 0 Or .Frames(frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                'If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                'If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                'If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                'If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum

                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
        End If
    Wend
    
    Close handle
    
    

    LoadGrhDataNew = True
Exit Function

ErrorHandler:
    LoadGrhDataNew = False
End Function

''
' Saves grh data using the old (and obsolete) file format. Shouldn't be used if possible.
' New format is valid with the new engine, included in Argentum Online 0.12.1
'
' @param    path    The complete path of the folde rin which to write the grh data file.
'                   If it existed it's deleted first.
'
' @return   True if the file was properly saved, False otherwise (data can't be stored in the old file format, use new one).

Public Function SaveGrhDataOld(ByVal path As String) As Boolean
    Dim handle
    Dim frame As Long
    Dim i As Long
    Dim tempint As Integer
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & GRH_DAT_FILE
    
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    MiCabecera.desc = OLD_FORMAT_HEADER
    
    'Write headers
    Put handle, , MiCabecera
    Put handle, , tempint
    Put handle, , tempint
    Put handle, , tempint
    Put handle, , tempint
    Put handle, , tempint
    
    'Store Grh List
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            'Index too big for this file format?
            If i > &H7FFF& Then
                Close handle
                Kill path
                Exit Function
            End If
            
            Put handle, , CInt(i)
            
            With GrhData(i)
                'Set number of frames
                Put handle, , .NumFrames
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For frame = 1 To .NumFrames
                        Put handle, , CInt(.Frames(frame))
                    Next frame
                    
                    Put handle, , CInt(.Speed * 0.018 / .NumFrames)
                Else
                    'Write in normal GRH data
                    Put handle, , CInt(.FileNum)
                    
                    Put handle, , .sX
                    
                    Put handle, , .sY
                        
                    Put handle, , .pixelWidth
                    
                    Put handle, , .pixelHeight
                End If
            End With
        End If
    Next i
    
    Close handle
    
    SaveGrhDataOld = True
End Function

''
' Saves grh data using the old (and obsolete) file format. Shouldn't be used if possible.
' New format is valid with the new engine, included in Argentum Online 0.12.1
'
' @param    path    The complete path of the folde rin which to write the grh data file.
'                   If it existed it's deleted first.
'
' @return   True if the file was properly saved, False otherwise.

Public Function SaveGrhDataNew(ByVal path As String) As Boolean
    Dim handle
    Dim frame As Long
    Dim i As Long
    Dim tempint As Integer
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & GRH_DAT_FILE
    
    
    handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As handle
    
    'Increment file version
    fileVersion = fileVersion + 1
    
    Put handle, , fileVersion
    
    Put handle, , CLng(UBound(GrhData()))
    
    'Store Grh List
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            Put handle, , i
            
            With GrhData(i)
                'Set number of frames
                Put handle, , .NumFrames
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For frame = 1 To .NumFrames
                        Put handle, , .Frames(frame)
                    Next frame
                    
                    Put handle, , .Speed
                Else
                    'Write in normal GRH data
                    Put handle, , .FileNum
                    
                    Put handle, , .sX
                    
                    Put handle, , .sY
                        
                    Put handle, , .pixelWidth
                    
                    Put handle, , .pixelHeight
                End If
            End With
        End If
    Next i
    
    Close handle
    
    SaveGrhDataNew = True
End Function
