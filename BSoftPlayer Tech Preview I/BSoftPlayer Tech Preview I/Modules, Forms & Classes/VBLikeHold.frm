VERSION 5.00
Begin VB.Form frmVBLikeHold 
   BorderStyle     =   0  'None
   Caption         =   "ffsdfhg"
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7500
   Begin VB.PictureBox m_Controls 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   7515
      TabIndex        =   2
      Top             =   0
      Width           =   7515
      Begin VB.PictureBox picTitlebar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         Picture         =   "VBLikeHold.frx":0000
         ScaleHeight     =   330
         ScaleWidth      =   7500
         TabIndex        =   3
         Top             =   0
         Width           =   7500
      End
   End
   Begin VB.CommandButton CmdNewCancel 
      Caption         =   "&Cancela"
      Height          =   345
      Left            =   5460
      TabIndex        =   1
      Top             =   960
      Width           =   1125
   End
   Begin VB.CommandButton cmdNewOpen 
      Caption         =   "&Opena"
      Default         =   -1  'True
      Height          =   345
      Left            =   5640
      TabIndex        =   0
      Top             =   660
      Width           =   1125
   End
   Begin VB.PictureBox cd 
      Height          =   480
      Left            =   6720
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   1920
      Width           =   1200
   End
End
Attribute VB_Name = "frmVBLikeHold"
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

Private Sub m_Controls_Click()

End Sub
