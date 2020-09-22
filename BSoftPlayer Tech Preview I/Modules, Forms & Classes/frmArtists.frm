VERSION 5.00
Begin VB.Form frmArtists 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmArtists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim SelectedItem As Integer
Dim OldSelectedItem As Integer
Public Rtrn As Integer

Private Sub Form_Load()
    OldSelectedItem = -1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        SelectedItem = Int(y / TextHeight("aaa"))
        Rtrn = SelectedItem
        Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If OldSelectedItem <> -1 Then
        ForeColor = vbWhite
        For i = SelectedItem * TextHeight("aaa") To SelectedItem * TextHeight("aaa") + TextHeight("aaa")
            Line (0, i)-(Me.Width, i)
        Next
        ForeColor = vbBlack
        CurrentY = SelectedItem * TextHeight("aaa")
        CurrentX = 40
        Print Items(SelectedItem)
        If ItmHasNums Then
            CurrentY = SelectedItem * TextHeight("aaa")
            CurrentX = Me.Width - 75 - Me.TextWidth("(" & ItmNumber(SelectedItem) & ")")
            Print "(" & ItmNumber(SelectedItem) & ")"
        End If
    End If
    SelectedItem = Int(y / TextHeight("aaa"))
    OldSelectedItem = SelectedItem
    ForeColor = vbHighlight
    For i = SelectedItem * TextHeight("aaa") To SelectedItem * TextHeight("aaa") + TextHeight("aaa")
        Line (0, i)-(Me.Width, i)
    Next
    ForeColor = vbHighlightText
    CurrentY = SelectedItem * TextHeight("aaa")
    CurrentX = 40
    Print Items(SelectedItem)
    If ItmHasNums Then
        CurrentY = SelectedItem * TextHeight("aaa")
        CurrentX = Me.Width - 75 - Me.TextWidth("(" & ItmNumber(SelectedItem) & ")")
        Print "(" & ItmNumber(SelectedItem) & ")"
    End If
End Sub
