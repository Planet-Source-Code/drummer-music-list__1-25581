VERSION 5.00
Begin VB.MDIForm frmMusicList 
   BackColor       =   &H8000000C&
   Caption         =   "Music List"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7905
   Icon            =   "frmMusicList.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowMusic 
         Caption         =   "&Music List"
      End
      Begin VB.Menu mnuWindowDir 
         Caption         =   "M&ake A New List"
      End
   End
End
Attribute VB_Name = "frmMusicList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed by Drummer
'Do not remove any comments from this program or refrences to Drummer so that i get full credit.
'You may use this code for ideas of you own but this idea was mine
Option Explicit



Private Sub MDIForm_Load()
    mintDate = "12/3/2001"
    mstrVersion = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    UnloadAll
End Sub

Private Sub mnuFileAbout_Click()
MsgBox "Music List          v " & mstrVersion & vbCrLf & "By Drummer" & vbCrLf & vbCrLf & "Programmed: " & mintDate, vbOKOnly, "About Music List"
End Sub

Private Sub mnuFileExit_Click()
    UnloadAll
End Sub

Private Sub mnuWindowDir_Click()
    frmDir.Show vbModeless
End Sub

Private Sub mnuWindowMusic_Click()
    frmMusic.Show vbModeless
End Sub
