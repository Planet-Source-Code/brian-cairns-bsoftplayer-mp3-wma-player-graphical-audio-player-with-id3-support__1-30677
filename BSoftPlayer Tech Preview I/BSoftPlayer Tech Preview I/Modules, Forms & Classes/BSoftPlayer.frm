VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "BSoftPlayer 1.0"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BSoftPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPL 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBEBEB&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   840
      ScaleHeight     =   285
      ScaleWidth      =   885
      TabIndex        =   24
      Top             =   2100
      Width           =   915
      Begin VB.Label lblPlaylist 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Playlist"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   0
         MouseIcon       =   "BSoftPlayer.frx":49E2
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   30
         Width           =   900
      End
   End
   Begin RichTextLib.RichTextBox rtbFile 
      Height          =   735
      Left            =   540
      TabIndex        =   23
      Top             =   1260
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"BSoftPlayer.frx":4B34
   End
   Begin VB.PictureBox picMutea 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   1740
      MouseIcon       =   "BSoftPlayer.frx":4BB9
      MousePointer    =   99  'Custom
      Picture         =   "BSoftPlayer.frx":4D0B
      ScaleHeight     =   300
      ScaleWidth      =   375
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMutea 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   1740
      MouseIcon       =   "BSoftPlayer.frx":533D
      MousePointer    =   99  'Custom
      Picture         =   "BSoftPlayer.frx":548F
      ScaleHeight     =   300
      ScaleWidth      =   375
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMute 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   60
      MouseIcon       =   "BSoftPlayer.frx":5AC1
      MousePointer    =   99  'Custom
      Picture         =   "BSoftPlayer.frx":5C13
      ScaleHeight     =   300
      ScaleWidth      =   375
      TabIndex        =   20
      Top             =   1020
      Width           =   375
   End
   Begin VB.PictureBox picVolThumb 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   180
      Picture         =   "BSoftPlayer.frx":6245
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   19
      Top             =   1425
      Width           =   240
   End
   Begin VB.PictureBox picVolume 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   210
      Picture         =   "BSoftPlayer.frx":6557
      ScaleHeight     =   585
      ScaleWidth      =   165
      TabIndex        =   18
      Top             =   1380
      Width           =   165
   End
   Begin VB.Timer tmrUpdateSeek 
      Interval        =   100
      Left            =   1500
      Top             =   1200
   End
   Begin VB.PictureBox picThumb 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   945
      Picture         =   "BSoftPlayer.frx":6B15
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   16
      Top             =   1785
      Width           =   255
   End
   Begin VB.PictureBox picSeek 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   900
      Picture         =   "BSoftPlayer.frx":6ECB
      ScaleHeight     =   165
      ScaleWidth      =   2610
      TabIndex        =   15
      Top             =   1860
      Width           =   2610
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EBEBEB&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   238
      TabIndex        =   13
      Top             =   2100
      Width           =   3600
      Begin VB.Label lblMore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "More"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1740
         MouseIcon       =   "BSoftPlayer.frx":8591
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   30
         Width           =   765
      End
      Begin VB.Label lblAdd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2700
         MouseIcon       =   "BSoftPlayer.frx":86E3
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   30
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblSettings 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         MouseIcon       =   "BSoftPlayer.frx":8835
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   30
         Width           =   735
      End
      Begin VB.Shape shpAdd 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   2520
         Top             =   -15
         Width           =   1515
      End
   End
   Begin VB.PictureBox picClosea 
      Height          =   615
      Left            =   600
      Picture         =   "BSoftPlayer.frx":8987
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   12
      Top             =   2460
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picCloseDown 
      Height          =   135
      Left            =   1500
      Picture         =   "BSoftPlayer.frx":8E99
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   11
      Top             =   2460
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picDropdownDown 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   -60
      Picture         =   "BSoftPlayer.frx":93AB
      ScaleHeight     =   315
      ScaleWidth      =   3600
      TabIndex        =   10
      Top             =   2460
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.PictureBox PicMinimizea 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      Picture         =   "BSoftPlayer.frx":CEFD
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   2460
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picMinimizeDown 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1500
      Picture         =   "BSoftPlayer.frx":D35F
      ScaleHeight     =   495
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picArtist 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      Picture         =   "BSoftPlayer.frx":D7C1
      ScaleHeight     =   315
      ScaleWidth      =   3600
      TabIndex        =   2
      Top             =   330
      Width           =   3600
      Begin VB.Label lblArtist 
         BackStyle       =   0  'Transparent
         Caption         =   "<< Artist >>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   45
         Width           =   3315
      End
   End
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      Picture         =   "BSoftPlayer.frx":11313
      ScaleHeight     =   315
      ScaleWidth      =   3600
      TabIndex        =   1
      Top             =   645
      Width           =   3600
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   45
         Width           =   3255
      End
   End
   Begin VB.PictureBox picTitlebar 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      Picture         =   "BSoftPlayer.frx":14E65
      ScaleHeight     =   330
      ScaleWidth      =   3600
      TabIndex        =   0
      Top             =   0
      Width           =   3600
      Begin VB.PictureBox picMinimize 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3105
         Picture         =   "BSoftPlayer.frx":18C87
         ScaleHeight     =   330
         ScaleWidth      =   225
         TabIndex        =   6
         Top             =   0
         Width           =   225
      End
      Begin VB.PictureBox picClose 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3330
         Picture         =   "BSoftPlayer.frx":190E9
         ScaleHeight     =   330
         ScaleWidth      =   270
         TabIndex        =   5
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.Image imgPreva 
      Height          =   330
      Index           =   1
      Left            =   0
      Picture         =   "BSoftPlayer.frx":195FB
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   0
      Y1              =   60
      Y2              =   152
   End
   Begin VB.Line Line5 
      X1              =   239
      X2              =   239
      Y1              =   60
      Y2              =   160
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   660
      TabIndex        =   17
      Top             =   1560
      Width           =   2595
   End
   Begin VB.Line Line4 
      X1              =   36
      X2              =   36
      Y1              =   64
      Y2              =   140
   End
   Begin VB.Line Line3 
      X1              =   40
      X2              =   244
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Image imgNexta 
      Height          =   330
      Index           =   1
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":19C6D
      Top             =   1860
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgNexta 
      Height          =   330
      Index           =   0
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":1A2DF
      Top             =   1740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgStopa 
      Height          =   330
      Index           =   2
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":1A951
      Top             =   2100
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgStopa 
      Height          =   330
      Index           =   1
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":1AFC3
      Top             =   1980
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgPausea 
      Height          =   330
      Index           =   2
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":1B635
      Top             =   2100
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgPausea 
      Height          =   330
      Index           =   1
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":1BCA7
      Top             =   1980
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgPausea 
      Height          =   330
      Index           =   0
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":1C319
      Top             =   1860
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgPlaya 
      Height          =   330
      Index           =   2
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":1C98B
      Top             =   2100
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgPlaya 
      Height          =   330
      Index           =   1
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":1CFFD
      Top             =   1980
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgPlaya 
      Height          =   330
      Index           =   0
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":1D66F
      Top             =   1860
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgStop 
      Height          =   330
      Left            =   1500
      Picture         =   "BSoftPlayer.frx":1DCE1
      Top             =   1020
      Width           =   345
   End
   Begin VB.Image imgPause 
      Height          =   330
      Left            =   1080
      Picture         =   "BSoftPlayer.frx":1E353
      Top             =   1020
      Width           =   345
   End
   Begin VB.Image imgPrev 
      Height          =   330
      Left            =   1980
      Picture         =   "BSoftPlayer.frx":1E9C5
      Top             =   1020
      Width           =   345
   End
   Begin VB.Image imgNext 
      Height          =   330
      Left            =   2400
      Picture         =   "BSoftPlayer.frx":1F037
      Top             =   1020
      Width           =   345
   End
   Begin VB.Image imgPlay 
      Height          =   330
      Left            =   660
      Picture         =   "BSoftPlayer.frx":1F6A9
      Top             =   1020
      Width           =   345
   End
   Begin MediaPlayerCtl.MediaPlayer wmPlayer 
      Height          =   675
      Left            =   3660
      TabIndex        =   9
      Top             =   1380
      Width           =   3015
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -130
      WindowlessVideo =   0   'False
   End
   Begin VB.Image imgStopa 
      Height          =   330
      Index           =   0
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":1FD1B
      Top             =   1860
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgPreva 
      Height          =   330
      Index           =   0
      Left            =   -1500
      Picture         =   "BSoftPlayer.frx":2038D
      Top             =   1740
      Visible         =   0   'False
      Width           =   345
   End
End
Attribute VB_Name = "frmMain"
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
Public CurrentArtist As Integer
Public CurrentTrack As Integer
Public PlayPending As Boolean
Dim Txa As Integer, DragNowa As Boolean
Dim Tyb, DragNowb As Boolean
Dim Muted As Boolean
Dim PlaylistShown As Boolean

Private Sub Form_Load()
    Dim i As Byte
    If Command <> "" Then
        If Len(Command) > 6 Then
            If LCase(Right(Command, 5)) = ".mp3""" Then
                shpAdd.Visible = True
                lblAdd.Visible = True
                frmNotInCol.Show vbModal, Me
                PlayPending = True
                wmPlayer.Open Mid(Command, 2, Len(Command) - 2)
            End If
        End If
    End If
    SetLayered Me.hWnd, True, 0
    Me.Show
    lblArtist.Caption = "Loading playlist..."
    Me.Refresh
    For i = 0 To 253 Step 1
        SetLayered Me.hWnd, True, i
    Next
    SetLayered Me.hWnd, False, 0
    LoadPlaylist
End Sub

Private Sub Image5_Click()

End Sub

Private Sub HScroll1_Change()
    wmPlayer.Volume = HScroll1.Value
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal And PlaylistShown Then frmPlaylist.WindowState = vbNormal
End Sub

Private Sub imgNext_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgNext.Picture = imgNexta(1).Picture
End Sub

Private Sub imgNext_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgNext.Picture = imgNexta(0).Picture
    NextTrack
End Sub

Private Sub imgPause_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If wmPlayer.PlayState = mpPlaying Then
        imgPause.Picture = imgPausea(1).Picture
        wmPlayer.Pause
    End If
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If wmPlayer.PlayState <> mpPlaying And wmPlayer.OpenState = 6 Then
        imgPlay.Picture = imgPlaya(1).Picture
        wmPlayer.Play
    End If
End Sub

Private Sub imgPrev_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgPrev.Picture = imgPreva(1).Picture
End Sub

Private Sub imgPrev_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgPrev.Picture = imgPreva(0).Picture
    PrevTrack
End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If wmPlayer.PlayState <> mpStopped Then
        imgStop.Picture = imgStopa(0).Picture
        imgStop.Refresh
        wmPlayer.Stop
        wmPlayer.CurrentPosition = 0
    End If
End Sub

Private Sub lblArtist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picArtist_MouseDown 0, 0, 0, 0
End Sub



Private Sub lblPlaylist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If PlaylistShown Then
        frmPlaylist.HideMe
        PlaylistShown = False
        picPL.BackColor = &HEBEBEB
        PLDockedBottom = False
        frmPlaylist.lnDock.Visible = False
        picPL.Height = 21
    Else
        picPL.BackColor = 16775133
        picPL.Height = 22
        PLDockedBottom = True
        frmPlaylist.lnDock.Visible = True
        frmPlaylist.ShowMe Me.top + Me.Height - Screen.TwipsPerPixelY, Me.left
        PlaylistShown = True
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picTitle_MouseDown 0, 0, 0, 0
End Sub

Private Sub picArtist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picArtist.Picture = picDropdownDown.Picture
    picArtist.Refresh
    Dim MaxWidth As Integer
    MaxWidth = Me.Width
    frmArtists.Height = (UBound(modParsePlaylist.Artists) + 1) * frmArtists.TextHeight("AAA") + (2 * Screen.TwipsPerPixelY)
    ReDim Items(UBound(modParsePlaylist.Artists))
    ReDim ItmNumber(UBound(modParsePlaylist.Artists))
    ItmHasNums = True
    For i = 0 To UBound(modParsePlaylist.Artists)
        frmArtists.CurrentX = 40
        frmArtists.CurrentY = i * frmArtists.TextHeight("aaa")
        frmArtists.Print modParsePlaylist.Artists(i).Name
        Items(i) = modParsePlaylist.Artists(i).Name
        ItmNumber(i) = UBound(modParsePlaylist.Artists(i).Songs) + 1
        
        If frmArtists.TextWidth(modParsePlaylist.Artists(i).Name) + 250 + frmArtists.TextWidth("(" & UBound(modParsePlaylist.Artists(i).Songs) + 1 & ")") > MaxWidth Then
            MaxWidth = frmArtists.TextWidth(modParsePlaylist.Artists(i).Name) + 250 + frmArtists.TextWidth("(" & UBound(modParsePlaylist.Artists(i).Songs) + 1 & ")")
        End If
    Next
    frmArtists.Width = MaxWidth + (2 * Screen.TwipsPerPixelX)
    frmArtists.top = Me.top + ((picArtist.top + picArtist.Height - 5) * Screen.TwipsPerPixelY)
    frmArtists.left = Me.left + (5 * Screen.TwipsPerPixelX)
    For i = 0 To UBound(modParsePlaylist.Artists)
        frmArtists.CurrentY = i * frmArtists.TextHeight("aaa")
        frmArtists.CurrentX = frmArtists.Width - 75 - frmArtists.TextWidth("(" & UBound(modParsePlaylist.Artists(i).Songs) + 1 & ")")
        frmArtists.Print "(" & UBound(modParsePlaylist.Artists(i).Songs) + 1 & ")"
    Next
    frmArtists.Show vbModal, Me
    If CurrentArtist <> frmArtists.Rtrn Then
        lblAdd.Visible = False
        shpAdd.Visible = False
        CurrentArtist = frmArtists.Rtrn
        lblArtist.Caption = modParsePlaylist.Artists(CurrentArtist).Name
        lblTitle.Caption = modParsePlaylist.Artists(CurrentArtist).Songs(0).Title
        CurrentTrack = 0
        wmPlayer.Open modParsePlaylist.Artists(CurrentArtist).Songs(CurrentTrack).URL
        PlayPending = True
    End If
    picArtist.Picture = picTitle.Picture
End Sub

Private Sub picClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picClose.Picture = picCloseDown.Picture
End Sub

Private Sub picClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Show
    Me.Refresh
    Dim i As Byte
    SetLayered Me.hWnd, True, 250
    Me.Refresh
    i = 250
    offset = wmPlayer.Volume
    For z = 1 To 250
        wmPlayer.Volume = -16 * (250 - i) + offset
        i = i - 1
        SetLayered Me.hWnd, True, i
        Me.Refresh
    Next
    End
End Sub

Private Sub picMinimize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picMinimize.Picture = picMinimizeDown.Picture
End Sub

Private Sub picMinimize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picMinimize.Picture = PicMinimizea.Picture
    Me.WindowState = vbMinimized
    frmPlaylist.WindowState = vbMinimized
End Sub

Private Sub picMute_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Muted Then
        picMute.Picture = picMutea(0).Picture
        wmPlayer.Mute = False
        Muted = False
    Else
        picMute.Picture = picMutea(1).Picture
        wmPlayer.Mute = True
        Muted = True
    End If
End Sub

Private Sub picThumb_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragNowa = True
    Txa = x
End Sub

Private Sub picThumb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If DragNowa Then
        NewLeft = picThumb.left + x - Txa
        If NewLeft < picSeek.left + 3 Then
            NewLeft = picSeek.left + 3
        End If
        If NewLeft > picSeek.Width + picSeek.left - 7 - picThumb.Width Then
            NewLeft = picSeek.Width + picSeek.left - 7 - picThumb.Width
        End If
        picThumb.left = NewLeft
    End If
End Sub

Private Sub picThumb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim offseti As Single
    DragNowa = False
    offseti = (picThumb.left - picSeek.left - 3) / (picSeek.Width - 10 - picThumb.Width)
    wmPlayer.CurrentPosition = Int(wmPlayer.Duration * offseti)
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If UBound(modParsePlaylist.Artists(CurrentArtist).Songs) = 0 Then
        Beep
        Exit Sub
    End If
    picTitle.Picture = picDropdownDown.Picture
    picTitle.Refresh
    Dim MaxWidth As Integer
    MaxWidth = Me.Width
    ItmHasNums = False
    frmArtists.Height = (UBound(modParsePlaylist.Artists(CurrentArtist).Songs) + 1) * frmArtists.TextHeight("AAA") + (2 * Screen.TwipsPerPixelY)
    ReDim Items(UBound(modParsePlaylist.Artists(CurrentArtist).Songs))
    For i = 0 To UBound(modParsePlaylist.Artists(CurrentArtist).Songs)
        frmArtists.CurrentX = 40
        frmArtists.Print modParsePlaylist.Artists(CurrentArtist).Songs(i).Title
        If frmArtists.TextWidth(modParsePlaylist.Artists(CurrentArtist).Songs(i).Title) + 40 > MaxWidth Then
            MaxWidth = frmArtists.TextWidth(modParsePlaylist.Artists(CurrentArtist).Songs(i).Title) + 40
        End If
        Items(i) = modParsePlaylist.Artists(CurrentArtist).Songs(i).Title
    Next
    frmArtists.Width = MaxWidth + (2 * Screen.TwipsPerPixelX)
    frmArtists.top = Me.top + ((picTitle.top + picTitle.Height - 5) * Screen.TwipsPerPixelY)
    frmArtists.left = Me.left + (5 * Screen.TwipsPerPixelX)
    frmArtists.Show vbModal, Me
    If CurrentTrack <> frmArtists.Rtrn Then
        lblAdd.Visible = False
        shpAdd.Visible = False
        CurrentTrack = frmArtists.Rtrn
        lblTitle.Caption = modParsePlaylist.Artists(CurrentArtist).Songs(CurrentTrack).Title
        wmPlayer.Open modParsePlaylist.Artists(CurrentArtist).Songs(CurrentTrack).URL
        PlayPending = True
    End If
    picTitle.Picture = picArtist.Picture
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
        If PLDockedBottom Then
            frmPlaylist.top = Me.top + Me.Height - Screen.TwipsPerPixelY
            frmPlaylist.left = Me.left
        End If
    End If
End Sub

Private Sub picTitlebar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragNow = False
End Sub


Sub LoadPlaylist()
    ReDim Tracks(0)
    ReDim Artists(0)
    modParsePlaylist.NumTracks = 0
    Dim Playlist As String
    Dim TempStr As String
    If Dir(App.Path & "\playlist.bsp") = "" Then
        imgNext.Enabled = False
        imgPrev.Enabled = False
        picTitle.Enabled = False
        picArtist.Enabled = False
        lblArtist.Caption = "No tracks loaded."
        lblTitle.Caption = "Click ""playlist"" to add tracks."
        modParsePlaylist.NumTracks = -1
        Exit Sub
    End If
    rtbFile.LoadFile App.Path & "\playlist.bsp"
    Playlist = rtbFile.Text
    lblArtist.Caption = "Parsing playlist..."
    lblArtist.Refresh
    modParsePlaylist.ParsePlaylist Playlist
    lblArtist.Refresh
    lblArtist.Caption = "Sorting artists..."
    modParsePlaylist.SortArtists
'    For i = 0 To UBound(modParsePlaylist.Artists)
'        Debug.Print modParsePlaylist.Artists(i).Name & ":"
'        For j = 0 To UBound(modParsePlaylist.Artists(i).Songs)
'            Debug.Print vbTab & modParsePlaylist.Artists(i).Songs(j).Title
'        Next
'    Next
    If modParsePlaylist.NumTracks = -1 Then
        imgNext.Enabled = False
        imgPrev.Enabled = False
        picTitle.Enabled = False
        picArtist.Enabled = False
        lblArtist.Caption = "No tracks loaded."
        Exit Sub
    End If
    If PlayPending = False Then
        lblArtist.Caption = modParsePlaylist.Artists(0).Name
        lblTitle.Caption = modParsePlaylist.Artists(0).Songs(0).Title
        PlayPending = True
        wmPlayer.Open modParsePlaylist.Artists(0).Songs(0).URL
    Else
        Dim Track As modID3.Info
        Track = modID3.GetID3(wmPlayer.Filename)
        lblArtist.Caption = Track.sArtist
        lblTitle.Caption = Track.sTitle
    End If
    picArtist.Enabled = True
    picTitle.Enabled = True
    imgNext.Enabled = True
    imgPrev.Enabled = True
    imgPause.Enabled = True
    imgPlay.Enabled = True
    imgStop.Enabled = True
End Sub

Private Sub ScrVolume_Change()
End Sub





Private Sub Picture2_Click()

End Sub

Private Sub picVolThumb_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Tyb = y
    DragNowb = True
End Sub

Private Sub picVolThumb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim offsetp As Single
    If DragNowb Then
        newPos = picVolThumb.top + y - Tyb
        If newPos < picVolume.top + 3 Then
            newPos = picVolume.top + 3
        End If
        If newPos > picVolume.top + picVolume.Height - 5 - picVolThumb.Height Then
            newPos = picVolume.top + picVolume.Height - 5 - picVolThumb.Height
        End If
        picVolThumb.top = newPos
    End If
    offsetp = (picVolThumb.top - picVolume.top - 3) / (picVolume.Height - 8 - picVolThumb.Height)
    wmPlayer.Volume = -4000 * offsetp
End Sub

Private Sub picVolThumb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragNowb = False
End Sub

Private Sub tmrUpdateSeek_Timer()
    Dim tm As Integer, tt As Integer, tp As Single, offset As Integer
    tm = Int(wmPlayer.CurrentPosition)
    tt = Int(wmPlayer.Duration)
    If tm <> -1 Then
        lblTime.Caption = FormatTime(tm) & " / " & FormatTime(tt)
        tp = tm / tt
        offset = Int((picSeek.Width - 10 - picThumb.Width) * tp)
        If Not DragNowa Then picThumb.left = offset + picSeek.left + 3
    Else
        lblTime.Caption = "..."
    End If
End Sub

Function FormatTime(tm As Integer) As String
    If tm = -1 Then
        timetext = "..."
    ElseIf tm < 60 Then
        timetext = ":" & Format(tm, "00")
    ElseIf tm < 3600 Then
        TimeMins = Int(tm / 60)
        TimeSecs = tm - TimeMins * 60
        timetext = Format(TimeMins, "00") & ":" & Format(TimeSecs, "00")
    Else
        timeHrs = Int(tm / 3600)
        TimeMins = Int(tm / 60) - (timeHrs * 60)
        TimeSecs = tm - TimeMins * 60 - timeHrs * 3600
        timetext = timeHrs & ":" & Format(TimeMins, "00") & ":" & Format(TimeSecs, "00")
    End If
    FormatTime = timetext
End Function

Private Sub wmPlayer_EndOfStream(ByVal Result As Long)
    NextTrack
End Sub

Sub NextTrack()
    lblAdd.Visible = False
    shpAdd.Visible = False
    If CurrentTrack <> UBound(modParsePlaylist.Artists(CurrentArtist).Songs) Then
        CurrentTrack = CurrentTrack + 1
    Else
        CurrentTrack = 0
        If CurrentArtist <> UBound(modParsePlaylist.Artists) Then
            CurrentArtist = CurrentArtist + 1
        Else
            CurrentArtist = 0
        End If
        lblArtist.Caption = modParsePlaylist.Artists(CurrentArtist).Name
    End If
    lblTitle.Caption = modParsePlaylist.Artists(CurrentArtist).Songs(CurrentTrack).Title
    wmPlayer.Open (modParsePlaylist.Artists(CurrentArtist).Songs(CurrentTrack).URL)
    PlayPending = True
End Sub

Sub PrevTrack()
    lblAdd.Visible = False
    shpAdd.Visible = False
    If CurrentTrack <> 0 Then
        CurrentTrack = CurrentTrack - 1
    Else
        If CurrentArtist <> 0 Then
            CurrentArtist = CurrentArtist - 1
            CurrentTrack = UBound(modParsePlaylist.Artists(CurrentArtist).Songs)
        Else
            CurrentArtist = UBound(modParsePlaylist.Artists)
            CurrentTrack = UBound(modParsePlaylist.Artists(CurrentArtist).Songs)
        End If
        lblArtist.Caption = modParsePlaylist.Artists(CurrentArtist).Name
    End If
    lblTitle.Caption = modParsePlaylist.Artists(CurrentArtist).Songs(CurrentTrack).Title
    wmPlayer.Open (modParsePlaylist.Artists(CurrentArtist).Songs(CurrentTrack).URL)
    PlayPending = True
End Sub

Private Sub wmPlayer_Error()
    Debug.Print wmPlayer.ErrorCode
    Select Case wmPlayer.ErrorCode
        Case -2147024894 'File not found
            Beep
            NextTrack
    End Select
End Sub

Private Sub wmPlayer_OpenStateChange(ByVal OldState As Long, ByVal NewState As Long)
    If NewState = 6 Then
        imgPlay.Picture = imgPlaya(0).Picture
        If PlayPending Then
            wmPlayer.Play
            PlayPending = False
        End If
    Else
        imgPlay.Picture = imgPlaya(2).Picture
    End If
End Sub

Private Sub wmPlayer_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    Select Case NewState
        Case 0
            imgPause.Picture = imgPausea(2).Picture
            imgPlay.Picture = imgPlaya(0).Picture
            imgStop.Picture = imgStopa(1).Picture
        Case 1
            imgPause.Picture = imgPausea(1).Picture
            imgPlay.Picture = imgPlaya(0).Picture
            imgStop.Picture = imgStopa(0).Picture
        Case 2
            imgPause.Picture = imgPausea(0).Picture
            imgPlay.Picture = imgPlaya(1).Picture
            imgStop.Picture = imgStopa(0).Picture
    End Select
End Sub
