VERSION 5.00
Begin VB.Form frmMusic 
   Caption         =   "Drummer's Music List"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6540
   Icon            =   "Music.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox lstSongs 
      Height          =   9225
      ItemData        =   "Music.frx":0442
      Left            =   480
      List            =   "Music.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   6015
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "Clear List Before Adding"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.ComboBox cboAlphabet 
      Height          =   5640
      ItemData        =   "Music.frx":0446
      Left            =   0
      List            =   "Music.frx":049B
      Style           =   1  'Simple Combo
      TabIndex        =   0
      ToolTipText     =   "To select alaphabetical list"
      Top             =   360
      Width           =   390
   End
   Begin VB.Label lblListCount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print List"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSeperator2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditRemove 
         Caption         =   "&Remove Line"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "C&lear List"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy List"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuGenre 
      Caption         =   "&Genre"
      Begin VB.Menu mnuData 
         Caption         =   "&1"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuData 
         Caption         =   "&2"
         Index           =   1
         Shortcut        =   {F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuData 
         Caption         =   "&3"
         Index           =   2
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuData 
         Caption         =   "&4"
         Index           =   3
         Shortcut        =   {F4}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuData 
         Caption         =   "&5"
         Index           =   4
         Shortcut        =   {F5}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuData 
         Caption         =   "&6"
         Index           =   5
         Shortcut        =   {F6}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuData 
         Caption         =   "&7"
         Index           =   6
         Shortcut        =   {F7}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuData 
         Caption         =   "&8"
         Index           =   7
         Shortcut        =   {F8}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuData 
         Caption         =   "&9"
         Index           =   8
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuData 
         Caption         =   "&10"
         Index           =   9
         Shortcut        =   {F11}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGenreListAll 
         Caption         =   "&List All"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuFunction 
      Caption         =   "F&unction"
      Begin VB.Menu mnuEditNumber 
         Caption         =   "&Number Items on List"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete List"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed by Drummer
'Do not remove comments from this code.
'You may use code within this for ideas of you own.
'This was my own idea and did not copy anyones idea, just bits of code.
Option Explicit

'declare to execute
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
        
Dim mintRecordCount As Long
Dim mintTable As String
Dim mstrAlpha As String
Dim mstrWidthSetting As String
Dim mstrHeightSetting As String
Dim WithEvents rsSongs As ADODB.Recordset
Attribute rsSongs.VB_VarHelpID = -1
Dim WithEvents rsSettings As ADODB.Recordset
Attribute rsSettings.VB_VarHelpID = -1
Dim FSO As New FileSystemObject, SngLst As File

Private Sub cboAlphabet_Click()
    Dim intNum As Integer
    'change list by number
    If rsSongs.EOF = True And rsSongs.BOF = True Then
    MsgBox "Database is empty", , "Error"
    Else
    ClearList
    If cboAlphabet.ListIndex = 0 Then
        intNum = 0
        Do Until intNum = 10
            mintTable = intNum
            NumberSearch
            intNum = intNum + 1
        Loop
    'change list by alphabet
    ElseIf cboAlphabet.ListIndex = 1 Then
        mstrAlpha = "a"
    ElseIf cboAlphabet.ListIndex = 2 Then
        mstrAlpha = "b"
    ElseIf cboAlphabet.ListIndex = 3 Then
        mstrAlpha = "c"
    ElseIf cboAlphabet.ListIndex = 4 Then
        mstrAlpha = "d"
    ElseIf cboAlphabet.ListIndex = 5 Then
        mstrAlpha = "e"
    ElseIf cboAlphabet.ListIndex = 6 Then
        mstrAlpha = "f"
    ElseIf cboAlphabet.ListIndex = 7 Then
        mstrAlpha = "g"
    ElseIf cboAlphabet.ListIndex = 8 Then
        mstrAlpha = "h"
    ElseIf cboAlphabet.ListIndex = 9 Then
        mstrAlpha = "i"
    ElseIf cboAlphabet.ListIndex = 10 Then
        mstrAlpha = "j"
    ElseIf cboAlphabet.ListIndex = 11 Then
        mstrAlpha = "k"
    ElseIf cboAlphabet.ListIndex = 12 Then
        mstrAlpha = "l"
    ElseIf cboAlphabet.ListIndex = 13 Then
        mstrAlpha = "m"
    ElseIf cboAlphabet.ListIndex = 14 Then
        mstrAlpha = "n"
    ElseIf cboAlphabet.ListIndex = 15 Then
        mstrAlpha = "o"
    ElseIf cboAlphabet.ListIndex = 16 Then
        mstrAlpha = "p"
    ElseIf cboAlphabet.ListIndex = 17 Then
        mstrAlpha = "q"
    ElseIf cboAlphabet.ListIndex = 18 Then
        mstrAlpha = "r"
    ElseIf cboAlphabet.ListIndex = 19 Then
        mstrAlpha = "s"
    ElseIf cboAlphabet.ListIndex = 20 Then
        mstrAlpha = "t"
    ElseIf cboAlphabet.ListIndex = 21 Then
        mstrAlpha = "u"
    ElseIf cboAlphabet.ListIndex = 22 Then
        mstrAlpha = "v"
    ElseIf cboAlphabet.ListIndex = 23 Then
        mstrAlpha = "w"
    ElseIf cboAlphabet.ListIndex = 24 Then
        mstrAlpha = "x"
    ElseIf cboAlphabet.ListIndex = 25 Then
        mstrAlpha = "y"
    ElseIf cboAlphabet.ListIndex = 26 Then
        mstrAlpha = "z"
    End If
    
    If cboAlphabet.ListIndex > 0 Then
        RecordSearchTwo
    End If
    lblListCount.Caption = lstSongs.ListCount & " songs in list"
    End If
End Sub

Private Sub cboAlphabet_KeyPress(KeyAscii As Integer)
        If KeyAscii > 31 And KeyAscii < 127 Then
            ClearList
            mstrAlpha = Chr(KeyAscii)
            RecordSearchTwo
            lblListCount.Caption = lstSongs.ListCount & " songs in list"
        End If
        cboAlphabet.Text = ""
End Sub

Private Sub Form_Activate()
    'setup connections for ado controls
    Set rsSongs = New ADODB.Recordset
    Set rsSettings = New ADODB.Recordset
    mintDate = "12/3/2001"
    mstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    'check for songs.lst file
    Dim strTemp As String
    Dim strTemp2 As String
    strTemp = Dir(App.Path & "\" & mstrOpen)
    strTemp2 = Dir(App.Path & "\Settings.ini")
    If strTemp2 <> "" Then
          rsSettings.Open App.Path & "\Settings.ini"
    End If
    If strTemp = "" Then
        'file does not exist, do nothing to buttons
        frmMusic.Hide
        frmDir.Show vbModeless
        frmDir.lblNo.Visible = True
        mnuEditDelete.Visible = False
    Else
        'file exists, continue program
        rsSongs.Open App.Path & "\" & mstrOpen
        mnuEditDelete.Visible = True
        'clear genre menu and then add new info
        Dim intNum As Integer
        Dim strNum As String
        intNum = 0
        rsSongs.MoveFirst
        Do Until intNum = 10
            mnuData(intNum).Caption = ""
            mnuData(intNum).Visible = False
            intNum = intNum + 1
        Loop
        intNum = 0
        Do Until intNum = 10
            strNum = intNum
            If rsSongs.Fields(strNum) <> "" And rsSongs.Fields(strNum) <> " " Then
                mnuData(intNum).Caption = rsSongs.Fields(strNum)
                mnuData(intNum).Visible = True
            End If
            intNum = intNum + 1
        Loop
        cboAlphabet.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mnuEditNumber.Enabled = False
    mstrOpen = "Songs.lst"
End Sub

Private Sub Form_Resize()
On Error GoTo Error
    'make listbox just smaller than window
    Dim x As Integer
    Dim Y As Integer
    x = ScaleWidth - 550
    Y = ScaleHeight - 360
    lstSongs.Width = x
    lstSongs.Height = Y
    Exit Sub
Error:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadAll
End Sub

Private Sub lstSongs_DblClick()     'Execute file when double clicked
    Dim strDir As String
    Dim strFile As String
    Dim strSearch As String
    Dim intNum As Integer
    Dim strNum As String
    Dim blnDone As Boolean
    blnDone = False
    
    rsSettings.MoveFirst
    intNum = 0
    'assign name from list and find
    Do Until intNum = 10 Or blnDone = True
        strNum = intNum
        With rsSongs    'check if ends with .mp3
            strFile = lstSongs.Text
            strFile = Replace(strFile, "'", "_*")
            strSearch = "[" & strNum & "] Like '" & strFile & ".mp3'"
            .MoveFirst
            .Find strSearch
            If .EOF Then    'check if ends with .mp3.lnk
                strFile = lstSongs.Text
                strFile = Replace(strFile, "'", "_*")
                strSearch = "[" & strNum & "] Like '" & strFile & ".mp3.lnk'"
                .MoveFirst
                .Find strSearch
                If .EOF Then    'check for the rest
                    strFile = lstSongs.Text
                    strFile = Replace(strFile, "'", "_*")
                    strSearch = "[" & strNum & "] Like '" & strFile & "'"
                    .MoveFirst
                    .Find strSearch
                    If .EOF Then
                    Else
                        blnDone = True
                    End If
                Else
                    blnDone = True
                End If
            Else
                blnDone = True
            End If
        End With
        intNum = intNum + 1
    Loop
    If blnDone = True Then      'execute from file (the meat)
        strFile = rsSongs.Fields(strNum)
        strFile = Replace(strFile, "`", "'")
        strDir = rsSettings.Fields(intNum + 9)
        strDir = Replace(strDir, "`", "'")
        ShellExecute 0, "open", strDir & strFile, "", strDir, 1     'the magic code        refer to declare at top
    End If
    
    rsSongs.MoveFirst
End Sub

Private Sub lstSongs_KeyUp(KeyCode As Integer, Shift As Integer)    'execute on push of enter
    If KeyCode = 13 Then
        lstSongs_DblClick
    End If
End Sub


Private Sub mnuData_Click(Index As Integer)     'display in list according to the name the user input
    If rsSongs.EOF = True And rsSongs.BOF = True Then
    MsgBox "Database is empty", , "Error"
    Else
    ClearList
    mintTable = Index
    RecordSearch
    End If
End Sub

Private Sub mnuEditClear_Click()
    'clear the list
    
    If MsgBox("Do you REALLY want to clear the list?", vbOKCancel, "Sure?") = vbOK Then
        lstSongs.Clear
    End If
    lblListCount.Caption = lstSongs.ListCount & " songs in list"
    mnuEditNumber.Enabled = False
End Sub

Private Sub mnuEditCopy_Click()                 'copy to clipboard
    Dim intNum As Integer
    Dim strList As String
    Clipboard.Clear

    intNum = 0
    Do Until intNum = lstSongs.ListCount
        strList = strList & lstSongs.List(intNum) & vbCrLf
        intNum = intNum + 1
    Loop
    Clipboard.SetText strList
End Sub

Private Sub mnuEditDelete_Click()   'delete Database file, hide frm, show list creator frm
    Dim strDir As String
    strDir = App.Path & "\Songs.lst"
    If MsgBox("This will permanently remove the list from your hard drive." & vbCrLf _
    & "Do you want to continue?", vbOKCancel, "Sure?") = vbOK Then
        FSO.DeleteFile (strDir)
        Me.Hide
        frmDir.Show
        lstSongs.Clear
        rsSongs.Close
    End If
End Sub

Private Sub mnuEditFind_Click()
    frmSearch.Show
End Sub

Private Sub mnuEditNumber_Click()
    'add a number to each item on the list
    'user should only do once each time they add
    Dim intX As Integer
    intX = 1
    Dim strSong As String
    Dim int10, int100, int1000, int10000 As String
    
    'create 0's before each number depending on how many songs are in the list
    'the more songs in the list, the more 0's before the numbers
    If lstSongs.ListCount < 10 Then
        int10 = ""
        int100 = ""
        int1000 = ""
        int10000 = ""
    End If
    If lstSongs.ListCount > 9 Then
        int10 = "0"
        int100 = ""
        int1000 = ""
        int10000 = ""
    End If
    If lstSongs.ListCount > 99 Then
        int10 = "00"
        int100 = "0"
        int1000 = ""
        int10000 = ""
    End If
    If lstSongs.ListCount > 999 Then
        int10 = "000"
        int100 = "00"
        int1000 = "0"
        int10000 = ""
    End If
    If lstSongs.ListCount > 9999 Then
        int10 = "0000"
        int100 = "000"
        int1000 = "00"
        int10000 = "0"
    End If
    Do Until intX = lstSongs.ListCount + 1      'loop till end of listbox
        lstSongs.ListIndex = intX - 1           'select index
        strSong = lstSongs.Text                 'save string
        If intX < 10 Then                       'apply proper 0's
            strSong = int10 & intX & ".  " & strSong
        ElseIf intX < 100 Then
            strSong = int100 & intX & ".  " & strSong
        ElseIf intX < 1000 Then
            strSong = int1000 & intX & ".  " & strSong
        ElseIf intX < 10000 Then
            strSong = int10000 & intX & ".  " & strSong
        End If
        lstSongs.RemoveItem intX - 1            'remove previous instance of index
        lstSongs.AddItem strSong                'add new one
        intX = intX + 1
    Loop
    mnuEditNumber.Enabled = False
End Sub

Private Sub mnuEditRemove_Click()
    'remove selected item from list
    
    If lstSongs.ListCount = 0 Then
        MsgBox "Must add to the list first.", , "List Is Empty"
    Else
    If lstSongs.ListIndex > -1 Then
        lstSongs.RemoveItem lstSongs.ListIndex
    Else
        MsgBox "HEY!, Pick an item to remove first.", vbOKOnly & vbExclamation, "Oops"
    End If
    End If
    lblListCount.Caption = lstSongs.ListCount & " songs in list"
End Sub

Private Sub mnuFileAbout_Click()
    'display msgbox for about
    Dim strDate As String
    
    If rsSongs.EOF = True And rsSongs.BOF = True Then
    MsgBox "Music List          v " & mstrVersion & vbCrLf & "By Drummer" & vbCrLf & vbCrLf & "Programmed: " & mintDate, vbOKOnly, "About Music List"
    Else
    rsSongs.MoveFirst
    strDate = rsSongs.Fields("Date")
    MsgBox "Music List          v " & mstrVersion & vbCrLf & "By Drummer" & vbCrLf & vbCrLf & "Programmed: " & mintDate & vbCrLf & "Music List Dated: " & strDate, vbOKOnly, "About Music List"
    End If
End Sub


Private Sub lstSongs_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'right click short cut
    
    If Button = vbRightButton Then
        PopupMenu mnuGenre
    End If
End Sub

Private Sub mnuFileExit_Click()
    UnloadAll
End Sub

Private Sub mnuFileOpen_Click()
    frmOpen.Show vbModal
    Form_Activate
End Sub

Private Sub mnuFilePrint_Click()
    'Print what is in the list        hidden for now, use copy to clipboard
    
    Dim intIndex As Integer
    Dim intFinalValue As Integer
    intFinalValue = lstSongs.ListCount - 1
    
    If MsgBox("Do You Still Want To Print?" & vbCrLf & "This is not formatted and will look ugly", vbOKCancel + vbQuestion, "Sure?") = vbOK Then
        For intIndex = 0 To intFinalValue
            Printer.Print lstSongs.List(intIndex)
        Next intIndex
    End If
    Printer.EndDoc
End Sub

Private Sub mnuFileReturn_Click()
    frmMusic.Hide
End Sub

Private Sub ClearList()
    'clear list
    
    If chkClear.Value = Checked Then
        lstSongs.Clear
    End If
End Sub

Private Sub RecordSearch()          'add all songs from one Genre
    MousePointer = vbHourglass
    Dim strSong As String
    With rsSongs
        .MoveNext
        Do Until .EOF
            If .Fields(mintTable) <> "" Then
                strSong = .Fields(mintTable)
                If Right(strSong, 4) = ".mp3" Then
                    strSong = Left(strSong, Len(strSong) - 4)      'remove .mp3 extension
                ElseIf Right(strSong, 4) = ".lnk" Then
                    strSong = Left(strSong, Len(strSong) - 8)      'remove .lnk extension
                End If
                lstSongs.AddItem strSong
            End If
            .MoveNext
        Loop
        .MoveFirst
    End With
    lblListCount.Caption = lstSongs.ListCount & " songs in list"
    MousePointer = vbDefault
    mnuEditNumber.Enabled = True
End Sub

Private Sub RecordSearchTwo()       'add songs beginning with specified letter
    MousePointer = vbHourglass
    Dim strSong As String
    Dim intNum As Integer
    intNum = 0
    With rsSongs
        Do Until intNum = 10
            mintTable = intNum
            .MoveFirst
            .MoveNext
            Do Until .EOF
                If UCase(mstrAlpha) = Left(UCase(.Fields(mintTable)), 1) Then
                    strSong = .Fields(mintTable)
                    If Right(strSong, 4) = ".mp3" Then
                        strSong = Left(strSong, Len(strSong) - 4)      'remove .mp3 extension
                    ElseIf Right(strSong, 4) = ".lnk" Then
                        strSong = Left(strSong, Len(strSong) - 8)      'remove .lnk extension
                    End If
                    lstSongs.AddItem strSong
                End If
                .MoveNext
            Loop
            intNum = intNum + 1
        Loop
        .MoveFirst
    End With
    MousePointer = vbDefault
    mnuEditNumber.Enabled = True
End Sub

Private Sub NumberSearch()            'add songs beginning with any number
    MousePointer = vbHourglass
    Dim strSong As String
    With rsSongs
        .MoveNext
        Do Until .EOF
            If 9 > Left(UCase(.Fields(mintTable)), 1) Then
                strSong = .Fields(mintTable)
                If Right(strSong, 4) = ".mp3" Then
                    strSong = Left(strSong, Len(strSong) - 4)      'remove .mp3 extension
                ElseIf Right(strSong, 4) = ".lnk" Then
                    strSong = Left(strSong, Len(strSong) - 8)      'remove .lnk extension
                End If
                lstSongs.AddItem strSong
            End If
            .MoveNext
        Loop
        .MoveFirst
    End With
    MousePointer = vbDefault
    mnuEditNumber.Enabled = True
End Sub

Private Sub mnuGenreListAll_Click() 'Add the whole thing to the list
    MousePointer = vbHourglass
    Dim strSong As String
    Dim intNum As Integer
    ClearList
    intNum = 0
    With rsSongs
        Do Until intNum = 10
            .MoveNext
            Do Until .EOF
                If .Fields(intNum) <> "" Then
                    strSong = .Fields(intNum)
                    If Right(strSong, 4) = ".mp3" Then
                        strSong = Left(strSong, Len(strSong) - 4)      'remove .mp3 extension
                    ElseIf Right(strSong, 4) = ".lnk" Then
                        strSong = Left(strSong, Len(strSong) - 8)      'remove .lnk extension
                    End If
                    lstSongs.AddItem strSong
                End If
                .MoveNext
            Loop
            intNum = intNum + 1
            .MoveFirst
        Loop
    End With
    lblListCount.Caption = lstSongs.ListCount & " songs in list"
    MousePointer = vbDefault
    mnuEditNumber.Enabled = True
End Sub

Private Sub mnuWindowDir_Click()
    frmDir.Show vbModeless
End Sub

Private Sub mnuWindowMusic_Click()
    frmMusic.Show vbModeless
End Sub
