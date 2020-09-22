Attribute VB_Name = "modParsePlaylist"
'BSoftPlayer Audio Player Technology Preview
'Copyright (C) 2001, 2002 BSoft, Inc.
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

Public Tracks() As Track
Private TracksTemp() As String

Private TrackBeginLocation As Long
Private TrackEndLocation As Long

Private GenBeginLocation As Integer
Private GenEndLocation As Integer

Private i As Integer, j As Integer

Public Artists() As Artist
Dim NumArtists As Integer

Public Type Track
    Title As String
    Artist As String
    Album As String
    URL As String
End Type

Public Type Artist
    Name As String
    Songs() As Track
End Type



Dim Tags(0 To 3) As String
Dim TagsE(0 To 3) As String
Dim TempString As String

Public NumTracks As Long
Sub AddItem(URL As String, Title As String, Artist As String, Album As String)
    If modParsePlaylist.NumTracks = -1 Then
        ReDim Tracks(0)
    Else
        ReDim Preserve Tracks(UBound(Tracks) + 1)
    End If
    Tracks(UBound(Tracks)).URL = URL
    Tracks(UBound(Tracks)).Title = Title
    Tracks(UBound(Tracks)).Artist = Artist
    Tracks(UBound(Tracks)).Album = Album
    modParsePlaylist.NumTracks = 1
End Sub

Sub ReSortAndIndex()
    ReDim Artists(0)
    SortArtists
    If frmMain.lblArtist.Caption = "No tracks loaded." Then
        frmMain.CurrentArtist = 0
        frmMain.CurrentTrack = 0
        frmMain.lblArtist.Caption = modParsePlaylist.Artists(0).Name
        frmMain.lblTitle.Caption = modParsePlaylist.Artists(0).Songs(0).Title
    Else
        frmMain.CurrentArtist = FindArtist(frmMain.lblArtist.Caption)
        frmMain.CurrentTrack = FindTrack(frmMain.lblTitle.Caption, frmMain.CurrentArtist)
    End If
    frmPlaylist.lvMusic.ListItems.Clear
    frmPlaylist.LoadTracks
End Sub

Private Function CountNumTracks(XML As String) As Long
    Dim NumTracksTmp As Long
    Dim OldStart As Long
    OldStart = 1
    Do While OldStart <> 0
        OldStart = InStr(OldStart + 1, XML, "<Entry>", vbTextCompare)
        NumTracksTmp = NumTracksTmp + 1
    Loop
    CountNumTracks = NumTracksTmp - 2
End Function

Sub ParsePlaylistASX(XML As String)
' Determine number of tracks
    NumTracks = CountNumTracks(XML)
' Dimension arrays
    ReDim Tracks(NumTracks)
    ReDim TracksTemp(NumTracks)
' Identify the item tags (start)
    Tags(0) = "<Param Name = ""Artist"" Value = """
    Tags(1) = "<Param Name = ""Name"" Value = """
    Tags(2) = "<Param Name = ""SourceURL"" Value = """
    Tags(3) = "<Param Name = ""Album"" Value = """
' Identify the item tags (end)
    TagsE(0) = """ />"
    TagsE(1) = """ />"
    TagsE(2) = """ />"
    TagsE(3) = """ />"
' Parse the XML into individual tracks
    For i = 0 To NumTracks
        TrackBeginLocation = InStr(TrackEndLocation + 1, XML, "<Entry>", vbTextCompare)
        TrackEndLocation = InStr(TrackBeginLocation + 1, XML, "</Entry>", vbTextCompare)
        TracksTemp(i) = Mid$(XML, TrackBeginLocation + 8, TrackEndLocation - TrackBeginLocation - 8)
        frmMain.lblArtist.Caption = "Parsing Playlist... (" & Int(i / NumTracks * 100) & "%)"
        frmMain.lblArtist.Refresh
    Next
'Send the individual tracks to another subroutine for additional processing
    For i = 0 To NumTracks
        ParseTrack TracksTemp(i), i
    Next
End Sub

Private Sub ParseTrack(XML As String, Index As Integer)
'Parse the individual track attributes
    For j = 0 To 3
        GenBeginLocation = 0
        GenEndLocation = 0
        GenBeginLocation = InStr(GenEndLocation + 1, XML, Tags(j), vbTextCompare)
        If GenBeginLocation <> 0 Then
            GenEndLocation = InStr(GenBeginLocation + 1, XML, TagsE(j), vbTextCompare)
            TempString = Mid(XML, GenBeginLocation + Len(Tags(j)), GenEndLocation - GenBeginLocation - Len(Tags(j)))
            SetTrackAttribute Index, j, TempString 'Send the story attribute to another subroutine
        End If
    Next
End Sub

Private Sub SetTrackAttribute(Index As Integer, AttributeIndex As Integer, Data As String)
    Select Case AttributeIndex
        Case 0 'Artist
            Tracks(Index).Artist = Data
        Case 1 'Title
            Tracks(Index).Title = Data
        Case 2 'URL
            Tracks(Index).URL = Data
        Case 3 'Album
            Tracks(Index).Album = Data
    End Select
End Sub

Sub SortArtists()
    If NumTracks = -1 Then Exit Sub
    NumArtists = -1
    Dim UniqueArtist As Boolean
    For i = 0 To NumTracks
        UniqueArtist = True
        If NumArtists <> -1 Then
            For j = 0 To NumArtists
                If Tracks(i).Artist = Artists(j).Name Or (Tracks(i).Artist = "" And Artists(j).Name = "- No Artist -") Then
                    UniqueArtist = False
                    Exit For
                End If
            Next
        End If
        If UniqueArtist Then
            NumArtists = NumArtists + 1
            ReDim Preserve Artists(NumArtists)
            Artists(NumArtists).Name = Tracks(i).Artist
            If Artists(NumArtists).Name = "" Then
                Artists(NumArtists).Name = "- No Artist -"
            End If
            ReDim Artists(NumArtists).Songs(0)
            Artists(NumArtists).Songs(0) = Tracks(i)
        Else
            ReDim Preserve Artists(j).Songs(UBound(Artists(j).Songs) + 1)
            Artists(j).Songs(UBound(Artists(j).Songs)) = Tracks(i)
        End If
    Next
    QSort 0, UBound(Artists)
    For i = 0 To UBound(Artists)
        QSortTracks Artists(i), 0, UBound(Artists(i).Songs)
    Next
End Sub

Sub ParsePlaylist(Txt As String)
    Dim TxtArr As Variant
    Dim TmpArr As Variant
    
    TxtArr = Split(Txt, vbCrLf)
       
    TmpArr = Split(TxtArr(0), Seperator)
    NumTracks = UBound(TmpArr)
    If NumTracks = -1 Then Exit Sub
    ReDim Tracks(NumTracks)
    For i = 0 To NumTracks
        Tracks(i).Title = TmpArr(i)
    Next
    
    TmpArr = Split(TxtArr(1), Seperator)
    For i = 0 To NumTracks
        Tracks(i).Artist = TmpArr(i)
    Next
    
    TmpArr = Split(TxtArr(2), Seperator)
    For i = 0 To NumTracks
        Tracks(i).Album = TmpArr(i)
    Next
    
    TmpArr = Split(TxtArr(3), Seperator)
    For i = 0 To NumTracks
        Tracks(i).URL = TmpArr(i)
    Next
End Sub

Sub SavePlaylist()
    Dim TitleArray() As String
    Dim ArtistArray() As String
    Dim AlbumArray() As String
    Dim URLArray() As String
    
    ReDim TitleArray(UBound(modParsePlaylist.Tracks))
    ReDim ArtistArray(UBound(modParsePlaylist.Tracks))
    ReDim AlbumArray(UBound(modParsePlaylist.Tracks))
    ReDim URLArray(UBound(modParsePlaylist.Tracks))
    
    For i = 0 To UBound(modParsePlaylist.Tracks)
        TitleArray(i) = modParsePlaylist.Tracks(i).Title
        ArtistArray(i) = modParsePlaylist.Tracks(i).Artist
        AlbumArray(i) = modParsePlaylist.Tracks(i).Album
        URLArray(i) = modParsePlaylist.Tracks(i).URL
    Next
    frmMain.rtbFile.Text = Join(TitleArray, Seperator) & vbCrLf & Join(ArtistArray, Seperator) & vbCrLf & Join(AlbumArray, Seperator) & vbCrLf & Join(URLArray, Seperator)
    frmMain.rtbFile.SaveFile App.Path & "\playlist.bsp", rtfText
End Sub
