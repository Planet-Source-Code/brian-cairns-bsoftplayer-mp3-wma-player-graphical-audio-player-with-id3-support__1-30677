Attribute VB_Name = "modID3"
Private mvarFilename As String


Public Type Info
    sTitle As String
    sArtist As String
    sAlbum As String
    sComment As String
    sYear As String
    sGenre As String
    End Type
Private MP3Info As Info


Public Property Get Filename() As String
    Filename = mvarFilename
End Property


Private Function IsValidFile(ByVal sFilename) As Boolean
    Dim bOk As Boolean
    ' make sure file exists
    bOk = CBool(Dir(sFilename, vbHidden) <> "")
    
    Dim aExtensions, ext
    aExtensions = Array(".mp3", ".mp2", ".mp1", ".wma", ".mpg", ".asf", ".avi")
    Dim bOkayExtension As Boolean
    bOkayExtension = False


    If bOk Then


        For Each ext In aExtensions


            If InStr(1, sFilename, ext, vbTextCompare) > 0 Then
                bOkayExtension = True
            End If
        Next 'ext
    End If
    ValidFilename = (InStr(1, sFilename, "?") = 0)
    IsValidFile = bOk And bOkayExtension And ValidFilename
End Function


Public Function GetID3(ByVal sPassFilename As String) As Info
    Dim iFreefile As Integer
    Dim lFilePos As Long
    Dim sData As String * 128
    
    Dim sGenre() As String
    ' Genre
    Const sGenreMatrix As String = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
    ' Build the Genre array (VB6+ only)
    sGenre = Split(sGenreMatrix, "|")
    ' Store the filename (for "Get Filename"
    '     property)
    mvarFilename = sPassFilename
    ' Clear the info variables
    


    If Not IsValidFile(sPassFilename) Then ' bug fix
        Exit Function
    End If
    
    MP3Info.sTitle = ""
    MP3Info.sArtist = ""
    MP3Info.sAlbum = ""
    MP3Info.sYear = ""
    MP3Info.sComment = ""
    ' Ensure the MP3 file exists
    ' Retrieve the info data from the MP3
    iFreefile = FreeFile
    lFilePos = FileLen(mvarFilename) - 127


    If lFilePos > 0 Then ' bug fix
        Open mvarFilename For Binary As #iFreefile
        Get #iFreefile, lFilePos, sData
        Close #iFreefile
    End If
    
    ' Populate the info variables


    If left(sData, 3) = "TAG" Then
        MP3Info.sTitle = Mid(sData, 4, 30)
        MP3Info.sArtist = Mid(sData, 34, 30)
        MP3Info.sAlbum = Mid(sData, 64, 30)
        MP3Info.sYear = Mid(sData, 94, 4)
        MP3Info.sComment = Mid(sData, 98, 30)
        Dim lGenre
        lGenre = Asc(Mid(sData, 128, 1))


        If lGenre <= UBound(sGenre) Then
            MP3Info.sGenre = sGenre(lGenre)
        Else
            MP3Info.sGenre = ""
        End If
    Else
        MP3Info = GetInfo(mvarFilename)
    End If
If Trim(MP3Info.sTitle) = "" Then MP3Info = GetInfo(mvarFilename)
GetID3 = MP3Info
End Function
'' Try to get something meaningful out o
'     f the filename


Public Function GetInfo(ByVal sFilename) As Info
    Dim i As Info
    GetInfo = i
    Dim s As String
    s = sFilename
    Dim Dash1Pos As Integer, Dash2Pos As Integer
    

    If InStrRev(s, "\") > 0 Then 'it's a full path
        s = Mid(s, InStrRev(s, "\") + 1)
    End If
    
    'drop extension
    s = left(s, InStrRev(s, ".", , vbTextCompare) - 1)
    s = Replace(Trim(s), " ", " ")
    s = Trim(s)
    s = Replace(s, "_", " ")
    
    Select Case CountNumDashes(s)
        Case 0 'Typical "Title"
            i.sTitle = Trim(s)
        Case 1 'Typical "Artist - Title"
            Dash1Pos = InStr(s, "-")
            i.sArtist = Trim(left(s, Dash1Pos - 1))
            i.sTitle = Trim(Mid(s, Dash1Pos + 1))
        Case 2 'Typical "Artist - Album - Title"
    End Select
    GetInfo = i
End Function

Function CountNumDashes(stra As String) As Integer
    Dim NumDashes As Integer
    Dim CNumDashes As Integer
    NumDashes = 0
    Do While True
        NumDashes = InStr(NumDashes + 1, stra, "-")
        If NumDashes <> 0 Then
            CNumDashes = CNumDashes + 1
        Else
            Exit Do
        End If
    Loop
    CountNumDashes = CNumDashes
End Function

Private Function FixDir(sFullpath)
    Dim s1, s2
    s1 = Trim(left(sFullpath, InStrRev(sFullpath, "\") - 1))
    s2 = Trim(Mid(sFullpath, InStrRev(sFullpath, "\") + 1))
    FixDir = s1 & " - " & s2
End Function


