Attribute VB_Name = "Mejoras"
Option Explicit
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum
'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type
Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type
Public Type Grh
    grhindex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    loops As Integer
End Type
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceArma
    Arma(1 To 4) As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OFFSETX As Integer
    OFFSETY As Integer
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
    '[ANIM ATAK]
    ShieldAttack As Byte
End Type

Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData

Public MisCuerpos() As tIndiceCuerpo
Public MisCabezas() As tIndiceCabeza
Public MisCascos() As tIndiceCabeza
Public MisEscudos() As tIndiceArma
Public MisArmas() As tIndiceArma

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer
'Khalem
Public NumeroHechizos As Integer
Public Hechizos() As tHechizo
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer
Public NumCascos As Integer
Public NumEscudosAnims As Integer
Public MiCabecera As tCabecera


Public Type tHechizo
    Nombre As String
    WAV As Integer
    FXgrh As Integer
End Type

Public Type NpcData
    name As String
    Body As Integer
    Head As Integer
    Heading As Byte
    Agua As Boolean
    Comercia As Byte
    Alineacion As Byte
    NPCType As Byte
End Type
Public NumNPCs As Long
'Public NumNPCsHOST As Integer
Public NpcData() As NpcData

Public Type ObjData
    name As String 'Nombre del obj
    ObjType As Integer 'Tipo enum que determina cuales son las caract del obj
    grhindex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    Info As String
    Ropaje As Integer 'Indice del grafico del ropaje
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    Texto As String
End Type
Public NumOBJs As Integer
Public ObjData() As ObjData


Type SupData
    name As String
    Grh As Integer
    Width As Byte
    Height As Byte
    Block As Boolean
    Capa As Byte
End Type
Public MaxSup As Integer
Public SupData() As SupData
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

Sub CargarOtrosInit()
CargarAnimArmas
CargarAnimEscudos
CargarCabezas
CargarCascos
CargarCuerpos
CargarFxs
'Khalem
CargarHechizos

CargarIndicesSuperficie

CargarDats
End Sub
Sub CargarDats()
CargarIndicesOBJ
CargarIndicesNPC
End Sub
Sub GuardarOtrosInit()
GuardarArmas
GuardarEscudos
GuardarCabezas
GuardarCascos
GuardarCuerpos
GuardarFxs

GuardarIndicesSuperficie

End Sub
Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    Dim i As Integer
    Dim N As Integer
    N = FreeFile()
    Open Config.initPath & "Armas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumWeaponAnims
    

    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    ReDim MisArmas(1 To NumWeaponAnims) As tIndiceArma
    
    For loopc = 1 To NumWeaponAnims
        Get #N, , MisArmas(loopc)
    
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), MisArmas(loopc).Arma(1), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), MisArmas(loopc).Arma(2), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), MisArmas(loopc).Arma(3), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), MisArmas(loopc).Arma(4), 0
    Next loopc
    
    Close #N
End Sub
Sub GuardarArmas()
'on error Resume Next
    Dim N As Integer
    Dim i As Long
    Dim loopc As Long
    Dim arch As String
    
    N = FreeFile()
    Open Config.initPath & "Armas.ind" For Binary Access Write As #N
    
    
    Put #N, , MiCabecera

    Put #N, , NumWeaponAnims

    
    
    For loopc = 1 To NumWeaponAnims
        Put #N, , MisArmas(loopc)

    Next loopc
    
    Close #N
End Sub

Sub CargarAnimEscudos()
'on error Resume Next

    Dim loopc As Long
    Dim arch As String
    Dim N As Integer
    Dim i As Long

    N = FreeFile()
    Open Config.initPath & "Escudos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumEscudosAnims
    

    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    ReDim MisEscudos(1 To NumEscudosAnims) As tIndiceArma
    
    For loopc = 1 To NumEscudosAnims
        Get #N, , MisEscudos(loopc)
        
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), MisEscudos(loopc).Arma(1), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), MisEscudos(loopc).Arma(2), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), MisEscudos(loopc).Arma(3), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), MisEscudos(loopc).Arma(4), 0
    Next loopc
    
    Close #N
End Sub
Sub GuardarEscudos()
'on error Resume Next
    Dim N As Integer
    Dim i As Long
    Dim loopc As Long
    Dim arch As String
    
    N = FreeFile()
    Open Config.initPath & "Escudos.ind" For Binary Access Write As #N
    
    
    Put #N, , MiCabecera

    Put #N, , NumEscudosAnims

    
    
    For loopc = 1 To NumEscudosAnims
        Put #N, , MisEscudos(loopc)

    Next loopc
    
    Close #N
End Sub
Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long

    N = FreeFile()
    Open Config.initPath & "Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(1 To Numheads) As HeadData
    ReDim MisCabezas(1 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , MisCabezas(i)
        
        If MisCabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), MisCabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), MisCabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), MisCabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), MisCabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub
Sub GuardarCabezas()
    Dim N As Integer
    Dim i As Long

    N = FreeFile()
    Open Config.initPath & "Cabezas.ind" For Binary Access Write As #N
    
    'cabecera
    Put #N, , MiCabecera
    
    'num de cabezas
    Put #N, , Numheads
    
    For i = 1 To Numheads
        Put #N, , MisCabezas(i)

    Next i
    
    Close #N
End Sub

Sub CargarHechizos()
      
On Error GoTo Fallo
    If Dir(datPath & "hechizos.dat") = "" Then
        MsgBox "Falta el archivo 'hechizos.dat' en " & datPath, vbCritical
        End
    End If
    
        Dim Hechizo As Integer
    Dim Leer As New clsIniReader

    Call Leer.Initialize(datPath & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = Val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos
        With Hechizos(Hechizo)
            .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            .WAV = Val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = Val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
        End With
    Next Hechizo
    
    Set Leer = Nothing
    
    Exit Sub

Fallo:
MsgBox "Error al intentar cargar el Objteto " & Hechizo & " de hechizos.dat en " & datPath & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

 
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    

    
    
    N = FreeFile()
    Open Config.initPath & "Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(1 To NumCascos) As HeadData
    ReDim MisCascos(1 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , MisCascos(i)
        
        If MisCabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), MisCascos(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), MisCascos(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), MisCascos(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), MisCascos(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub
Sub GuardarCascos()
    Dim N As Integer
    Dim i As Long
    

    N = FreeFile()
    Open Config.initPath & "Cascos.ind" For Binary Access Write As #N
    
    'cabecera
    Put #N, , MiCabecera
    
    'num de cabezas
    Put #N, , NumCascos
    
    For i = 1 To NumCascos
        Put #N, , MisCascos(i)
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
 
    N = FreeFile()
    Open Config.initPath & "Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumBodies
    
    'Resize array
    ReDim BodyData(1 To NumBodies) As BodyData
    ReDim MisCuerpos(1 To NumBodies) As tIndiceCuerpo

    For i = 1 To NumBodies
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub
Sub GuardarCuerpos()
    Dim N As Integer
    Dim i As Long
    
    
    N = FreeFile()
    Open Config.initPath & "Personajes.ind" For Binary Access Write As #N
    
    'cabecera
    Put #N, , MiCabecera
    
    'num de cabezas
    Put #N, , NumBodies

    For i = 1 To NumBodies
        Put #N, , MisCuerpos(i)
    Next i
    
    Close #N
End Sub
Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    
    N = FreeFile()
    Open Config.initPath & "Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub
Sub GuardarFxs()
    Dim N As Integer
    Dim i As Long
    
    N = FreeFile()
    Open Config.initPath & "Fxs.ind" For Binary Access Write As #N
    
    'cabecera
    Put #N, , MiCabecera
    
    'num de cabezas
    Put #N, , NumFxs
    
    For i = 1 To NumFxs
        Put #N, , FxData(i)
    Next i
    
    Close #N
End Sub
Public Sub InitGrh(ByRef Grh As Grh, ByVal grhindex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.grhindex = grhindex
    If Grh.grhindex = 0 Then Exit Sub
    If Started = 2 Then
        If GrhData(Grh.grhindex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.grhindex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.loops = 999
    Else
        Grh.loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.grhindex).Speed
End Sub




Public Sub CargarIndicesOBJ()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
    If Dir(datPath & "OBJ.dat") = "" Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & datPath, vbCritical
        End
    End If
    Dim Obj As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(datPath & "OBJ.dat")

    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumOBJs) As ObjData
    For Obj = 1 To NumOBJs
        DoEvents
        ObjData(Obj).name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).grhindex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
        ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))

    Next Obj
    Exit Sub
Fallo:
MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & datPath & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub



Public Sub CargarIndicesNPC()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
On Error Resume Next
'on error GoTo Fallo
    If Dir(datPath & "NPCs.dat") = "" Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & datPath, vbCritical
        End
    End If
    'If FileExist(DirDats & "\NPCs-HOSTILES.dat", vbArchive) = False Then
    '    MsgBox "Falta el archivo 'NPCs-HOSTILES.dat' en " & DirDats, vbCritical
    '    End
    'End If
    Dim Trabajando As String
    Dim NPC As Integer
    Dim Leer As New clsIniReader

    Call Leer.Initialize(datPath & "NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    'Call Leer.Initialize(DirDats & "\NPCs-HOSTILES.dat")
    'NumNPCsHOST = Val(Leer.GetValue("INIT", "NumNPCs"))
    ReDim NpcData(1000) As NpcData
    Trabajando = "Dats\NPCs.dat"
    'Call Leer.Initialize(DirDats & "\NPCs.dat")
    'MsgBox "  "
    For NPC = 1 To NumNPCs
        NpcData(NPC).name = Leer.GetValue("NPC" & NPC, "Name")
        
        NpcData(NPC).Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
        NpcData(NPC).Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
        NpcData(NPC).Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
        NpcData(NPC).Agua = Val(Leer.GetValue("NPC" & NPC, "AguaValida")) = 1
        NpcData(NPC).Comercia = Val(Leer.GetValue("NPC" & NPC, "Comercia"))
        NpcData(NPC).Alineacion = Val(Leer.GetValue("NPC" & NPC, "Alineacion"))
        NpcData(NPC).NPCType = Val(Leer.GetValue("NPC" & NPC, "NPCType"))
    Next
    'MsgBox "  "
    'Trabajando = "Dats\NPCs-HOSTILES.dat"
    'Call Leer.Initialize(DirDats & "\NPCs-HOSTILES.dat")
    'For NPC = 1 To NumNPCsHOST
    '    NpcData(NPC + 499).name = Leer.GetValue("NPC" & (NPC + 499), "Name")
    '    NpcData(NPC + 499).Body = Val(Leer.GetValue("NPC" & (NPC + 499), "Body"))
    '    NpcData(NPC + 499).Head = Val(Leer.GetValue("NPC" & (NPC + 499), "Head"))
    '    NpcData(NPC + 499).Heading = Val(Leer.GetValue("NPC" & (NPC + 499), "Heading"))
    '    If LenB(NpcData(NPC + 499).name) <> 0 Then frmMain.lListado(2).AddItem NpcData(NPC + 499).name & " - #" & (NPC + 499)
    'Next NPC
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & datPath & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub


Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If Dir(indicePath) = "" Then
        MsgBox "Falta el archivo " & indicePath, vbCritical
        End
    End If
    Dim Leer As New clsIniReader
    Dim i As Integer
    Leer.Initialize indicePath
    MaxSup = Leer.GetValue("INIT", "Referencias")
    ReDim SupData(MaxSup) As SupData
    For i = 0 To MaxSup
        SupData(i).name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
    Next
    DoEvents
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de " & indicePath & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
End Sub

Public Sub GuardarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If Dir(indicePath) = "" Then
        MsgBox "Falta el archivo " & indicePath, vbCritical
        End
    End If
    Dim NF As Integer
    Dim i As Integer
    NF = FreeFile
    Open (indicePath) For Output As NF
        Print #NF, "[INIT]"
        Print #NF, "Referencias=" & MaxSup
        
        For i = 0 To MaxSup
            Print #NF, vbCrLf & "[REFERENCIA" & i & "]"
            Print #NF, "Nombre=" & SupData(i).name
            Print #NF, "GrhIndice=" & SupData(i).Grh
            Print #NF, "Ancho=" & SupData(i).Width
            Print #NF, "Alto=" & SupData(i).Height
            If SupData(i).Block Then Print #NF, "Bloquear=1"
            If SupData(i).Capa > 0 Then Print #NF, "Capa=" & SupData(i).Capa
        Next i
        
    Close NF
    DoEvents
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de " & indicePath & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
End Sub
