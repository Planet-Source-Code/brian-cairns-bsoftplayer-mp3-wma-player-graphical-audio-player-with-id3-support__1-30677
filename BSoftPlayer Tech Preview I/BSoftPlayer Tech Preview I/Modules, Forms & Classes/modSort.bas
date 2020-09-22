Attribute VB_Name = "modSort"
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

Option Explicit

Public Function QSort(lLbound As Long, lUbound As Long)
    Dim ArtistTmp As Artist
    Dim ArtBuffer As Artist
    Dim lngCurLow As Long
    Dim lngCurHigh As Long
    Dim lngCurMidpoint As Long
    
    lngCurLow = lLbound ' Start current low and high at actual low/high
    lngCurHigh = lUbound
    
    If lUbound <= lLbound Then Exit Function ' Error!
    lngCurMidpoint = (lLbound + lUbound) \ 2 ' Find the approx midpoint of the array
    
    ArtistTmp = Artists(lngCurMidpoint) ' Pick as a starting point (we are making
    ' an assumption that the data *might* be
    '
    ' in semi-sorted order already!
    
    Do While (lngCurLow <= lngCurHigh)


        Do While Artists(lngCurLow).Name < ArtistTmp.Name
            lngCurLow = lngCurLow + 1
            If lngCurLow = lUbound Then Exit Do
        Loop
        


        Do While ArtistTmp.Name < Artists(lngCurHigh).Name
            lngCurHigh = lngCurHigh - 1
            If lngCurHigh = lLbound Then Exit Do
        Loop


        If (lngCurLow <= lngCurHigh) Then ' if low is <= high then swap
            ArtBuffer = Artists(lngCurLow)
            Artists(lngCurLow) = Artists(lngCurHigh)
            Artists(lngCurHigh) = ArtBuffer
            '
            lngCurLow = lngCurLow + 1 ' CurLow++
            lngCurHigh = lngCurHigh - 1 ' CurLow--
        End If
        
    Loop
    


    If lLbound < lngCurHigh Then ' Recurse if necessary
        QSort lLbound, lngCurHigh
    End If
    


    If lngCurLow < lUbound Then ' Recurse if necessary
        QSort lngCurLow, lUbound
    End If
    
End Function

Public Function QSortTracks(ByRef Art As Artist, lLbound As Long, lUbound As Long)
    Dim TrackTmp As Track
    Dim TrackBuffer As Track
    Dim lngCurLow As Long
    Dim lngCurHigh As Long
    Dim lngCurMidpoint As Long
    
    lngCurLow = lLbound ' Start current low and high at actual low/high
    lngCurHigh = lUbound
    
    If lUbound <= lLbound Then Exit Function ' Error!
    lngCurMidpoint = (lLbound + lUbound) \ 2 ' Find the approx midpoint of the array
    
    TrackTmp = Art.Songs(lngCurMidpoint) ' Pick as a starting point (we are making
    ' an assumption that the data *might* be
    '
    ' in semi-sorted order already!
    
    Do While (lngCurLow <= lngCurHigh)


        Do While Art.Songs(lngCurLow).Title < TrackTmp.Title
            lngCurLow = lngCurLow + 1
            If lngCurLow = lUbound Then Exit Do
        Loop
        


        Do While TrackTmp.Title < Art.Songs(lngCurHigh).Title
            lngCurHigh = lngCurHigh - 1
            If lngCurHigh = lLbound Then Exit Do
        Loop


        If (lngCurLow <= lngCurHigh) Then ' if low is <= high then swap
            TrackBuffer = Art.Songs(lngCurLow)
            Art.Songs(lngCurLow) = Art.Songs(lngCurHigh)
            Art.Songs(lngCurHigh) = TrackBuffer
            '
            lngCurLow = lngCurLow + 1 ' CurLow++
            lngCurHigh = lngCurHigh - 1 ' CurLow--
        End If
        
    Loop
    


    If lLbound < lngCurHigh Then ' Recurse if necessary
        QSortTracks Art, lLbound, lngCurHigh
    End If
    


    If lngCurLow < lUbound Then ' Recurse if necessary
        QSortTracks Art, lngCurLow, lUbound
    End If
    
End Function

