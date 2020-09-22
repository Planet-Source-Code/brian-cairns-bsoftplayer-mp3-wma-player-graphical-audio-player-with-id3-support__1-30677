VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form frmOpen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Artists"
   ClientHeight    =   5430
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCheckAll 
      Cancel          =   -1  'True
      Caption         =   "Check &All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   4980
      Width           =   1515
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   4980
      Width           =   1035
   End
   Begin VB.FileListBox flMusic 
      Height          =   285
      Left            =   -1980
      TabIndex        =   3
      Top             =   1500
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   4980
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   1800
      Top             =   3240
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
            Picture         =   "frmOpen.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen.frx":015C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitlebar 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      Picture         =   "frmOpen.frx":02B8
      ScaleHeight     =   330
      ScaleWidth      =   10500
      TabIndex        =   0
      Top             =   0
      Width           =   10500
   End
   Begin CCRPFolderTV6.FolderTreeview FolderTreeview1 
      Height          =   5055
      Left            =   30
      TabIndex        =   2
      Top             =   345
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   8916
      Appearance      =   0
      IntegralHeight  =   0   'False
      VirtualFolders  =   0   'False
   End
   Begin VB.PictureBox picNone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   3180
      ScaleHeight     =   4515
      ScaleWidth      =   7275
      TabIndex        =   4
      Top             =   360
      Width           =   7275
      Begin VB.PictureBox picWait 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4515
         Left            =   0
         ScaleHeight     =   4515
         ScaleWidth      =   7275
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   7275
         Begin MSComctlLib.ProgressBar pbScan 
            Height          =   135
            Left            =   2040
            TabIndex        =   9
            Top             =   2220
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   238
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblHelp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Scanning folder for media..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   7
            Top             =   1920
            Width           =   2715
         End
      End
      Begin VB.Label lblHelp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No media found in this folder."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   0
         Left            =   2280
         TabIndex        =   5
         Top             =   1920
         Width           =   2715
      End
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   4515
      Left            =   3180
      TabIndex        =   8
      Top             =   360
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   7964
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Artist"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Album"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   208
      X2              =   700
      Y1              =   328
      Y2              =   328
   End
   Begin VB.Line Line1 
      X1              =   208
      X2              =   208
      Y1              =   20
      Y2              =   364
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      Height          =   5430
      Left            =   0
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "frmOpen"
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
Dim RealPath As String


Private Sub lvMusic_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvMusic.Sorted = True
    lvMusic.SortKey = ColumnHeader.Index - 1
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

Function FindTrack(TrackName As String, Artist As Integer)
    FindTrack = -1
    For i = 0 To UBound(modParsePlaylist.Artists(Artist).Songs)
        If TrackName = modParsePlaylist.Artists(Artist).Songs(i).Title Then
            FindTrack = i
            Exit Function
        End If
    Next
End Function

Function FindArtist(ArtistName As String)
    FindArtist = -1
    For i = 0 To UBound(modParsePlaylist.Artists)
        If ArtistName = modParsePlaylist.Artists(i).Name Then
            FindArtist = i
            Exit Function
        End If
    Next
End Function

Private Sub lvMusic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Beep
    
    End If
End Sub


Private Sub cmdAdd_Click()
    For i = 1 To lvFiles.ListItems.Count
        If lvFiles.ListItems(i).Checked Then
            AddItem RealPath & lvFiles.ListItems(i).Text, lvFiles.ListItems(i).SubItems(1), lvFiles.ListItems(i).SubItems(2), lvFiles.ListItems(i).SubItems(3)
        End If
    Next
    SavePlaylist
    frmMain.LoadPlaylist
    frmPlaylist.lvMusic.ListItems.Clear
    frmPlaylist.LoadTracks
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub



Private Sub cmdCheckAll_Click()
    For i = 1 To lvFiles.ListItems.Count
        lvFiles.ListItems(i).Checked = True
    Next
    cmdAdd.Enabled = True
End Sub

Private Sub FolderTreeview1_SelectionChange(Folder As CCRPFolderTV6.Folder, PreChange As Boolean, Cancel As Boolean)
    cmdAdd.Enabled = False
    picWait.Visible = True
    picWait.Refresh
    pbScan.Value = 0
    Dim Extension As String
    Dim Filename As String
    Dim ItemFound As Boolean
    Dim FileInfo As modID3.Info
    If Mid(Folder.FullPath, 3, 1) <> "\" Then
        picWait.Visible = False
        picNone.Visible = True
        Exit Sub
    End If
    flMusic.Path = Folder.FullPath
    If Right(Folder.FullPath, 1) = "\" Then
        RealPath = flMusic.Path
    Else
        RealPath = flMusic.Path & "\"
    End If
    
    lvFiles.ListItems.Clear
    If flMusic.ListCount > 0 Then
        pbScan.Max = flMusic.ListCount
    End If
    For i = 0 To flMusic.ListCount
        pbScan.Value = i
        Filename = flMusic.List(i)
        Extension = LCase(Right(Filename, 4))
        If Extension = ".mp3" Or Extension = ".wma" Then
            lvFiles.ListItems.Add , , Filename
            ItemFound = True
            FileInfo = modID3.GetID3(RealPath & Filename)
            lvFiles.ListItems(lvFiles.ListItems.Count).SubItems(1) = FileInfo.sTitle
            lvFiles.ListItems(lvFiles.ListItems.Count).SubItems(2) = FileInfo.sArtist
            lvFiles.ListItems(lvFiles.ListItems.Count).SubItems(3) = FileInfo.sAlbum
        End If
    Next
    picNone.Visible = Not ItemFound
    picWait.Visible = False
End Sub

Private Sub Form_Load()
    'lhWnd = SendMessageByLong(lvFiles.hWnd, LVM_GETHEADER, 0, 0)
    'If (lhWnd <> 0) Then
        'S = GetWindowLong(lhWnd, GWL_STYLE)
        'lS = lS And Not HDS_BUTTONS
        'SetWindowLong lhWnd, GWL_STYLE, lS
    'End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lvFiles_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    cmdAdd.Enabled = False
    For i = 1 To lvFiles.ListItems.Count
        If lvFiles.ListItems(i).Checked Then
            cmdAdd.Enabled = True
        End If
    Next
End Sub

Private Sub picTitlebar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Tx = x
    Ty = y
    DragNow = True
End Sub

Private Sub picTitlebar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If DragNow Then
        Me.top = Me.top + y - Ty
        Me.left = Me.left + x - Tx
    End If
End Sub

Private Sub picTitlebar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragNow = False
End Sub

Private Sub picWait_Click()

End Sub
