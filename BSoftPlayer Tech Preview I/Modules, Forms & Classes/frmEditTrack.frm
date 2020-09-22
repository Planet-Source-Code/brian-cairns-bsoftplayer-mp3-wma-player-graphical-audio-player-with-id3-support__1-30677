VERSION 5.00
Begin VB.Form frmEditTrack 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Edit Track"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOK 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3780
      Picture         =   "frmEditTrack.frx":0000
      ScaleHeight     =   330
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   3090
      Width           =   1035
   End
   Begin VB.PictureBox picCancel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   4860
      Picture         =   "frmEditTrack.frx":1222
      ScaleHeight     =   330
      ScaleWidth      =   1035
      TabIndex        =   10
      Top             =   3090
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "Will Work for Bandwidth"
      Top             =   2040
      Width           =   4515
   End
   Begin VB.TextBox txtArtist 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "The Broadband"
      Top             =   1680
      Width           =   4515
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "Mr. File"
      Top             =   1320
      Width           =   4515
   End
   Begin VB.PictureBox picInfo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   0
      Picture         =   "frmEditTrack.frx":2444
      ScaleHeight     =   810
      ScaleWidth      =   6000
      TabIndex        =   1
      Top             =   330
      Width           =   6000
      Begin VB.Label lblFiletype 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ".mp3 file"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   420
         Width           =   4995
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmEditTrack.frx":121A6
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblFilename 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MrFile.mp3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   180
         Width           =   4995
      End
   End
   Begin VB.PictureBox picTitlebar 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      Picture         =   "frmEditTrack.frx":124B0
      ScaleHeight     =   330
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F0F0F0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      Height          =   510
      Left            =   0
      Top             =   3000
      Width           =   6000
   End
   Begin VB.Label lblHelp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   180
      TabIndex        =   9
      Top             =   1380
      Width           =   795
   End
   Begin VB.Label lblHelp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Artist:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   8
      Top             =   1740
      Width           =   795
   End
   Begin VB.Label lblHelp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Album:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   2100
      Width           =   795
   End
   Begin VB.Shape Shape1 
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmEditTrack"
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

Private Sub Label1_Click()

End Sub

