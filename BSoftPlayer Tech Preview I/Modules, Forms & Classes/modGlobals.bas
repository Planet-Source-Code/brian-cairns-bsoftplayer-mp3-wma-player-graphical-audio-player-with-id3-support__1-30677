Attribute VB_Name = "modGlobals"
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

Global Items() As String
Global ItmNumber() As Integer
Global ItmHasNums As Boolean
Global Const Seperator = "<SEP>"
Global PLDockedBottom As Boolean

Public Function FindArtist(ArtistName As String)
    FindArtist = -1
    For i = 0 To UBound(modParsePlaylist.Artists)
        If ArtistName = modParsePlaylist.Artists(i).Name Then
            FindArtist = i
            Exit Function
        End If
    Next
End Function

Public Function FindTrack(TrackName As String, Artist As Integer)
    FindTrack = -1
    For i = 0 To UBound(modParsePlaylist.Artists(Artist).Songs)
        If TrackName = modParsePlaylist.Artists(Artist).Songs(i).Title Then
            FindTrack = i
            Exit Function
        End If
    Next
End Function
