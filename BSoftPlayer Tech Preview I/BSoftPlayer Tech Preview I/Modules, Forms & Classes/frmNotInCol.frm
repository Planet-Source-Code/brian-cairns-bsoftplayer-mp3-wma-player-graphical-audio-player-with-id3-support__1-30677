VERSION 5.00
Begin VB.Form frmNotInCol 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   188
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOKPic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   120
      Picture         =   "frmNotInCol.frx":0000
      ScaleHeight     =   330
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   1020
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picOKPic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   120
      Picture         =   "frmNotInCol.frx":1222
      ScaleHeight     =   330
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picOK 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3345
      Picture         =   "frmNotInCol.frx":2444
      ScaleHeight     =   330
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   2385
      Width           =   1035
   End
   Begin VB.PictureBox picTitlebar 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      Picture         =   "frmNotInCol.frx":3666
      ScaleHeight     =   330
      ScaleWidth      =   4500
      TabIndex        =   4
      Top             =   0
      Width           =   4500
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Don't tell me this again"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2415
      UseMaskColor    =   -1  'True
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   160
      Y2              =   4
   End
   Begin VB.Line Line1 
      X1              =   299
      X2              =   299
      Y1              =   152
      Y2              =   20
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   540
      Left            =   0
      Top             =   2280
      Width           =   4500
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To add the track to your collection, click ""add"" in the lower corner of the main window."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   1620
      Width           =   3915
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "This track is not in your collection."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   2955
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The track will disappear from the playlist editor and main window as soon as you choose a new track or close the player."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   180
      TabIndex        =   0
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "frmNotInCol"
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

Dim Tx As Integer, Ty As Integer
Dim DragNow As Boolean
Dim OKOver As Boolean, OKCheck As Boolean


Private Sub picOK_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picOK.Picture = picOKPic(1).Picture
    OKOver = True
    OKCheck = True
End Sub

Private Sub picOK_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If OKCheck Then
        If x > 0 And x < picOK.ScaleWidth And y > 0 And y < picOK.ScaleHeight Then
            If OKOver = False Then
                OKOver = True
                picOK.Picture = picOKPic(1).Picture
            End If
        Else
            If OKOver = True Then
                OKOver = False
                picOK.Picture = picOKPic(0).Picture
            End If
        End If
    End If
End Sub

Private Sub picOK_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picOK.Picture = picOKPic(0).Picture
    OKOver = False
    OKCheck = False
    If x > 0 And x < picOK.ScaleWidth Then
        If y > 0 And y < picOK.ScaleHeight Then
            Me.Hide
        End If
    End If
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


Private Sub cmdClose_Click()
    Me.Hide
End Sub

