Attribute VB_Name = "modMain"
'Programmed by Drummer
'Do not remove comments from this code.
'You may use code within this for ideas of you own.
'This was my own idea and did not copy anyones idea, just bits of code.
Option Explicit
    Public mintDate As Date
    Public mstrVersion As String
    Public mintTXTNum As Integer
    Public mstrOpen As String
    
Sub Main()
    'startup procedure
    
    Load frmMusic
    Load frmDir
    frmDir.Hide
    frmMusic.Show
End Sub

Public Sub UnloadAll()
    'unload all forms
    'Call coolClose(frmMusicList, 15)
    Dim eachform As Form
    For Each eachform In Forms
        Unload eachform
    Next
End Sub

Public Function coolClose(FormClose As Form, speed As Integer)
Do Until FormClose.Height <= 405
    DoEvents
    FormClose.Height = FormClose.Height - speed * 9
    FormClose.Top = FormClose.Top + speed * 5
Loop
Do Until FormClose.Width <= 1680
    DoEvents
    FormClose.Width = FormClose.Width - speed * 9
    FormClose.Left = FormClose.Left + speed * 5
Loop
Unload FormClose
End Function

