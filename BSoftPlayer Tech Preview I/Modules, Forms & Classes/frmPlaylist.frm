VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPlaylist 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "BSoftPlayer - Playlist"
   ClientHeight    =   5250
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   Picture         =   "frmPlaylist.frx":0000
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTaba 
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   3360
      Picture         =   "frmPlaylist.frx":0512
      ScaleHeight     =   375
      ScaleWidth      =   1695
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.PictureBox picTaba 
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2760
      Picture         =   "frmPlaylist.frx":2688
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   15
      Top             =   2460
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox picTabSelected 
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   1620
      Picture         =   "frmPlaylist.frx":3CAA
      ScaleHeight     =   375
      ScaleWidth      =   1695
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.PictureBox picTabSelected 
      AutoSize        =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   1620
      Picture         =   "frmPlaylist.frx":5E20
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   13
      Top             =   2460
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox picCloseDown 
      Height          =   15
      Left            =   5460
      Picture         =   "frmPlaylist.frx":7442
      ScaleHeight     =   15
      ScaleWidth      =   735
      TabIndex        =   11
      Top             =   5340
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picClosea 
      Height          =   915
      Left            =   5520
      Picture         =   "frmPlaylist.frx":7954
      ScaleHeight     =   855
      ScaleWidth      =   675
      TabIndex        =   10
      Top             =   5340
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   3480
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlaylist.frx":7E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlaylist.frx":7FC2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTabpanel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   0
      Left            =   60
      ScaleHeight     =   4455
      ScaleWidth      =   7395
      TabIndex        =   4
      Top             =   780
      Width           =   7395
      Begin VB.PictureBox picDeletea 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   2
         Left            =   1380
         Picture         =   "frmPlaylist.frx":811E
         ScaleHeight     =   330
         ScaleWidth      =   1290
         TabIndex        =   20
         Top             =   2940
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.PictureBox picDeletea 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   1
         Left            =   1380
         Picture         =   "frmPlaylist.frx":97B8
         ScaleHeight     =   330
         ScaleWidth      =   1290
         TabIndex        =   19
         Top             =   3300
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.PictureBox picDeletea 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   0
         Left            =   1380
         Picture         =   "frmPlaylist.frx":AE52
         ScaleHeight     =   330
         ScaleWidth      =   1290
         TabIndex        =   18
         Top             =   3660
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.PictureBox picDelete 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1380
         Picture         =   "frmPlaylist.frx":C4EC
         ScaleHeight     =   330
         ScaleWidth      =   1290
         TabIndex        =   17
         Top             =   4080
         Width           =   1290
      End
      Begin VB.PictureBox PicAdda 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   1
         Left            =   0
         Picture         =   "frmPlaylist.frx":DB86
         ScaleHeight     =   330
         ScaleWidth      =   1290
         TabIndex        =   8
         Top             =   3300
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.PictureBox PicAdda 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   0
         Left            =   0
         Picture         =   "frmPlaylist.frx":F220
         ScaleHeight     =   330
         ScaleWidth      =   1290
         TabIndex        =   7
         Top             =   3660
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.PictureBox picAdd 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         Picture         =   "frmPlaylist.frx":108BA
         ScaleHeight     =   330
         ScaleWidth      =   1290
         TabIndex        =   6
         Top             =   4080
         Width           =   1290
      End
      Begin MSComctlLib.ListView lvMusic 
         Height          =   4020
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   7091
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Artist"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Album"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   0
         X2              =   7380
         Y1              =   4020
         Y2              =   4020
      End
   End
   Begin VB.PictureBox picTab 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   1080
      Picture         =   "frmPlaylist.frx":11F54
      ScaleHeight     =   375
      ScaleWidth      =   1695
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picTab 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   15
      Picture         =   "frmPlaylist.frx":140CA
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   360
      Width           =   1110
   End
   Begin VB.PictureBox picTitlebar 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      Picture         =   "frmPlaylist.frx":156EC
      ScaleHeight     =   330
      ScaleWidth      =   7500
      TabIndex        =   0
      Top             =   0
      Width           =   7500
      Begin VB.PictureBox picClose 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   7230
         Picture         =   "frmPlaylist.frx":1D816
         ScaleHeight     =   330
         ScaleWidth      =   270
         TabIndex        =   9
         Top             =   0
         Width           =   270
      End
      Begin VB.Line lnDock 
         BorderColor     =   &H00FFF7DD&
         Visible         =   0   'False
         X1              =   840
         X2              =   1755
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picTabBg 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      Picture         =   "frmPlaylist.frx":1DD28
      ScaleHeight     =   375
      ScaleWidth      =   7500
      TabIndex        =   2
      Top             =   360
      Width           =   7500
   End
   Begin VB.PictureBox picTabpanel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   1
      Left            =   60
      ScaleHeight     =   4455
      ScaleWidth      =   7395
      TabIndex        =   12
      Top             =   780
      Visible         =   0   'False
      Width           =   7395
   End
   Begin VB.Shape Shape1 
      Height          =   5250
      Left            =   0
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmPlaylist"
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

Dim Tx As Integer, Ty As Integer, DragNow As Boolean
Option Compare Text

Private Sub Form_Load()
    LoadTracks
    CustomizeListview
End Sub

Sub ShowMe(top, left)
    Me.top = top
    Me.left = left
    Dim i As Byte
    SetLayered Me.hwnd, True, 0
    Me.Show vbModeless, frmMain
    Me.Refresh
    For i = 0 To 200 Step 4
        SetLayered Me.hwnd, True, i
    Next
    'SetLayered Me.hWnd, False, 0
End Sub
Sub LoadTracks()
    If modParsePlaylist.NumTracks = -1 Then Exit Sub
    For i = 0 To UBound(modParsePlaylist.Tracks)
        If Right(modParsePlaylist.Tracks(i).URL, 3) Like "wma" Then
            lvMusic.ListItems.Add , "K" & Str(i), modParsePlaylist.Tracks(i).Title, , 1
        Else
            lvMusic.ListItems.Add , "K" & Str(i), modParsePlaylist.Tracks(i).Title, , 2
        End If
        lvMusic.ListItems(lvMusic.ListItems.Count).SubItems(1) = modParsePlaylist.Tracks(i).Artist
        lvMusic.ListItems(lvMusic.ListItems.Count).SubItems(2) = modParsePlaylist.Tracks(i).Album
    Next
End Sub

Sub CustomizeListview()
'    lhWnd = SendMessageByLong(lvMusic.hWnd, LVM_GETHEADER, 0, 0)
'    If (lhWnd <> 0) Then
'        lS = GetWindowLong(lhWnd, GWL_STYLE)
'        lS = lS And Not HDS_BUTTONS
'        SetWindowLong lhWnd, GWL_STYLE, lS
'    End If
    lvMusic.Sorted = True
    lvMusic.SortKey = 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then frmMain.WindowState = vbNormal
End Sub

Private Sub lvMusic_DblClick()
    If Not IsNull(lvMusic.SelectedItem) Then
        frmMain.wmPlayer.Open modParsePlaylist.Tracks(Val(Mid(lvMusic.SelectedItem.Key, 3))).URL
        frmMain.CurrentArtist = FindArtist(modParsePlaylist.Tracks(Val(Mid(lvMusic.SelectedItem.Key, 3))).Artist)
        frmMain.CurrentTrack = FindTrack(modParsePlaylist.Tracks(Val(Mid(lvMusic.SelectedItem.Key, 3))).Title, frmMain.CurrentArtist)
        frmMain.lblArtist.Caption = modParsePlaylist.Artists(frmMain.CurrentArtist).Name
        frmMain.lblTitle.Caption = modParsePlaylist.Artists(frmMain.CurrentArtist).Songs(frmMain.CurrentTrack).Title
        frmMain.PlayPending = True
    End If
End Sub


Sub HideMe()
    Me.Show
    Me.Refresh
    Dim i As Byte
    'SetLayered Me.hWnd, True, 230
    Me.Refresh
    i = 200
    For z = 1 To 50
        i = i - 4
        SetLayered Me.hwnd, True, i
        'Me.Refresh
    Next
    Me.Hide
End Sub


Private Sub lvMusic_ItemClick(ByVal Item As MSComctlLib.ListItem)
    picDelete.Picture = picDeletea(0).Picture
End Sub

Private Sub lvMusic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Beep
    
    End If
End Sub

Private Sub picAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picAdd.Picture = PicAdda(1).Picture
End Sub

Private Sub picAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picAdd.Picture = PicAdda(0).Picture
                  frmOpen.Show vbModal, Me
End Sub

Private Sub picDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDelete.Picture = picDeletea(1).Picture
End Sub

Private Sub picDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDelete.Picture = picDeletea(0).Picture
End Sub

Private Sub picTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To picTabpanel.UBound
        picTabpanel(i).Visible = False
        picTab(i).Picture = picTaba(i).Picture
    Next
    picTabpanel(Index).Visible = True
    picTab(Index).Picture = picTabSelected(Index).Picture
End Sub

Private Sub picTitlebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Tx = X
    Ty = Y
    DragNow = True
End Sub

Private Sub picTitlebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewTop As Integer, NewLeft As Integer
    If DragNow Then
        NewLeft = Me.left + X - Tx
        NewTop = Me.top + Y - Ty
        If NewTop - 100 < frmMain.top + frmMain.Height And NewTop + 100 > frmMain.top + frmMain.Height Then
            NewTop = frmMain.top + frmMain.Height - Screen.TwipsPerPixelY
            If NewLeft - 100 < frmMain.left And NewLeft + 100 > frmMain.left Then
                NewLeft = frmMain.left
                If frmMain.picPL.BackColor <> 16775133 Then frmMain.picPL.BackColor = 16775133
                If frmMain.picPL.Height = 21 Then frmMain.picPL.Height = 22
                PLDockedBottom = True
                If lnDock.Visible = False Then lnDock.Visible = True
            Else
                If frmMain.picPL.BackColor <> 16775133 Then frmMain.picPL.BackColor = 16775133
                PLDockedBottom = False
                If lnDock.Visible = True Then lnDock.Visible = False
                If frmMain.picPL.Height = 22 Then frmMain.picPL.Height = 21
            End If
        Else
            If frmMain.picPL.BackColor <> 16775133 Then frmMain.picPL.BackColor = 16775133
            PLDockedBottom = False
            If lnDock.Visible = True Then lnDock.Visible = False
            If frmMain.picPL.Height = 22 Then frmMain.picPL.Height = 21
        End If
        Me.top = NewTop
        Me.left = NewLeft
    End If
End Sub

Private Sub picTitlebar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragNow = False
End Sub

Sub ShowOpen()
End Sub
