VERSION 5.00
Begin VB.Form frmDir 
   Caption         =   "List Maker"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7455
   Icon            =   "frmDir.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkListDir 
      Caption         =   "List Folders"
      Height          =   255
      Left            =   5280
      TabIndex        =   31
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkSubDir 
      Caption         =   "Search Subdirectories"
      Height          =   255
      Left            =   5280
      TabIndex        =   30
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdRestart2 
      Caption         =   "&Restart"
      Height          =   375
      Left            =   6360
      TabIndex        =   28
      ToolTipText     =   "To start over."
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK2 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      ToolTipText     =   "Click Me"
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "&Restart"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      ToolTipText     =   "To start over."
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      ToolTipText     =   "Click Me"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Type in information according to Directions."
      Top             =   3240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Type in information according to Directions."
      Top             =   2880
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Type in information according to Directions."
      Top             =   2520
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Type in information according to Directions."
      Top             =   2160
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Type in information according to Directions."
      Top             =   1800
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Type in information according to Directions."
      Top             =   1440
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Type in information according to Directions."
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Type in information according to Directions."
      Top             =   720
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Type in information according to Directions."
      Top             =   360
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Type in information according to Directions."
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblNo 
      Caption         =   "There is no database.  Please create one."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblFull 
      Caption         =   "Database is full.  To create a new list click ""Edit"", ""Delete List""."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lbl3 
      Caption         =   "Enter the directory each list will refer to.  This may take a minute after pressing OK."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   8
      Left            =   3840
      TabIndex        =   23
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   9
      Left            =   3840
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl2 
      Caption         =   "Enter name of each list as they are to appear in music list program."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lbl1 
      Caption         =   "To create a list of music enter number of Lists.  Enter only a number."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      TabIndex        =   10
      Top             =   0
      Width           =   1815
   End
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
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete List"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed by Drummer
'Do not remove comments from this code.
'You may use code within this for ideas of you own.
'This was my own idea and did not copy anyones idea, just bits of code.
Option Explicit

Dim MyFile, MyPath, MyName, strDir(0 To 9) As String
Dim blnBoo As Boolean
Dim blnSettings As Boolean
Dim WithEvents rsSongs As ADODB.Recordset
Attribute rsSongs.VB_VarHelpID = -1
Dim WithEvents rsSettings As ADODB.Recordset
Attribute rsSettings.VB_VarHelpID = -1
Dim mintTotal As Integer
Dim strDate As String
Dim strTextBoxNum As String
Dim FSO As New FileSystemObject, SngLst As File

Private Sub cmdBrowse_Click()
    Load frmBrowse
    frmBrowse.Show vbModal
End Sub

Private Sub cmdOK_Click()
    Dim intNum As Integer
    intNum = 0
    
    'first time OK is pushed
    'request number of genre's
    
    'verify txt is not empty
    Do Until intNum = 10
        If txtName(intNum).Visible = True Then
            If txtName(intNum).Text = "" Then
            MsgBox "Box is empty", , "Empty!"
            Exit Sub
            End If
        End If
        intNum = intNum + 1
    Loop
    
    'record list number to minttable
    If blnBoo = False Then
        If IsNumeric(txtName(0).Text) Then
        mintTotal = txtName(0).Text
            If txtName(0).Text > 0 Then
                If txtName(0).Text > 10 Then
                    MsgBox "10 is the max allowed", , "Too Many"
                    txtName(0).Text = ""
                Else
                    intNum = 1
                    Do Until intNum = mintTotal
                        txtName(intNum).Visible = True      'unhide txt boxes by number entered
                    intNum = intNum + 1
                    Loop
                    txtName(0).Text = ""
                    'read settings for name and input into txt boxes
                    If blnSettings = False Then
                        intNum = 0
                        Do Until intNum = mintTotal
                            txtName(intNum) = rsSettings.Fields(intNum)
                            intNum = intNum + 1
                        Loop
                    End If
                'setup for next button push
                txtName(0).SetFocus
                cmdRestart.Visible = True
                lbl1.Visible = False
                lblNo.Visible = False
                lbl2.Visible = True
                blnBoo = True
                End If
                Exit Sub
            Else
                MsgBox "Data is a number too low", , "Enter Higher Number"
            End If
        Else
            MsgBox "Data is not a number.", , "Enter Number!"
        End If
    End If
    
    'second time OK is pushed
    'set txt's as labels for each genre and clear txt boxes for directories
    If blnBoo = True Then
        intNum = 0
        Do Until intNum = 10
            If txtName(intNum).Text = "" Then
                txtName(intNum).Text = " "
            End If
            intNum = intNum + 1
        Loop
        
        'create fields for database file
        intNum = 0
        Do Until intNum = 10
            rsSongs.Fields.Append intNum, adVarChar, 100
            intNum = intNum + 1
        Loop
        'add date
        rsSongs.Fields.Append "Date", adVarChar, 12
        rsSongs.Open App.Path & "\Songs.lst"
        rsSongs.AddNew
        
        'add genre header name
        intNum = 0
        Do Until intNum = mintTotal
            rsSongs.Fields(intNum) = txtName(intNum)
            intNum = intNum + 1
        Loop
        rsSongs.Fields("Date") = strDate
        rsSongs.Update
        
        'clear window
        lbl1.Visible = False
        lbl2.Visible = False
        cmdOK.Visible = False
        cmdRestart.Visible = False
        
        'move txtbox text to lbl caption
        intNum = 0
        Do Until intNum = mintTotal
            lblName(intNum).Caption = txtName(intNum).Text
            strDir(intNum) = txtName(intNum)
        intNum = intNum + 1
        Loop
        
        'clear txtbox
        intNum = 0
        Do Until intNum = 10
            txtName(intNum).Text = ""
        intNum = intNum + 1
        Loop
        txtName(0).Text = "D:\Music\Example Folder\"
        txtName_GotFocus (0)
        
        'read settings for dir
        If blnSettings = False Then
            intNum = 0
            Do Until intNum = mintTotal
                txtName(intNum) = rsSettings.Fields(intNum + 10)
                intNum = intNum + 1
            Loop
        End If
        NewWindow
    End If
    
End Sub

'set up for cmdOK2 button
Private Sub NewWindow()
    Dim intLBL As Integer
    intLBL = 0
    'make objects visible or invisible
    cmdOK2.Visible = True
    cmdBrowse.Visible = True
    chkSubDir.Visible = True
    chkListDir.Visible = True
    cmdRestart2.Visible = True
    lbl3.Visible = True
    Do Until intLBL = mintTotal
        lblName(intLBL).Visible = True
    intLBL = intLBL + 1
    Loop
    txtName(0).SetFocus
    cmdOK.Default = False
    cmdOK2.Default = True
End Sub

'read txt's and lbl's,   verify data and input into database list file
Private Sub cmdOK2_Click()
    MousePointer = vbHourglass
    Dim intNum As Integer
    Dim strSQL As String
    Dim blnEmpty As Boolean
    Dim strDir(), strName() As String
    Dim intDirCount, intDirNum As Integer
    intDirCount = 1
    blnEmpty = False
    ReDim strDir(intDirCount)
    ReDim strName(intDirCount)
    
    'verify txt is not empty
    Do Until intNum = 10
        If txtName(intNum).Visible = True Then
            If txtName(intNum).Text = "" Then
            MsgBox "Box is empty", , "Empty!"
            Exit Sub
            End If
        End If
        intNum = intNum + 1
    Loop
    
    'put a \ after dir if not present
    intNum = 0
    Do Until intNum = mintTotal
    If Right(txtName(intNum).Text, 1) <> "\" Then
        txtName(intNum).Text = txtName(intNum).Text & "\"
    End If
    intNum = intNum + 1
    Loop
    
    'set components invisible
    intNum = 0
    Do Until intNum = 10
        txtName(intNum).Visible = False
        lblName(intNum).Visible = False
        intNum = intNum + 1
    Loop
    cmdOK2.Visible = False
    cmdBrowse.Visible = False
    chkSubDir.Visible = False
    chkListDir.Visible = False
    cmdRestart2.Visible = False
    lbl3.Visible = False
    
    ' put directories into database file
    intNum = 0
    Do Until intNum = mintTotal     'do once for each list
        intDirNum = 0
        intDirCount = 1
        strDir(intDirNum) = txtName(intNum).Text
        Do Until intDirNum = intDirCount             'for recursing directories
            MyPath = strDir(intDirNum) & strName(intDirNum) & "\"      ' Set the path.
            MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
            rsSongs.MoveFirst
            rsSongs.MoveNext
            Do While MyName <> ""               ' Start the loop.
                ' Ignore the current directory and the encompassing directory.
                If MyName <> "." And MyName <> ".." Then
                    If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then    'check if its dir
                        strDir(intDirCount) = MyPath     'add path to temp dir
                        strName(intDirCount) = MyName    'add file to temp dir
                        If chkSubDir = Checked Then
                            intDirCount = intDirCount + 1         'increment if subdir chkbox checked
                        End If
                        ReDim Preserve strDir(intDirCount)   'set undefined string to a number
                        ReDim Preserve strName(intDirCount)  'set undefined string to a number
                        If chkListDir = Checked Then
                            MyName = Replace(MyName, "'", "`")          'replace ' to ` for executing reasons (see frmmusiclist)
                            rsSongs.AddNew
                            rsSongs.Fields(intNum) = MyName
                            rsSongs.Update
                            rsSongs.MoveNext
                        End If
                    Else
                        MyName = Replace(MyName, "'", "`")          'replace ' to ` for executing reasons (see frmmusiclist)
                        rsSongs.AddNew
                        rsSongs.Fields(intNum) = MyName
                        rsSongs.Update
                        rsSongs.MoveNext
                    End If
                End If
                MyName = Dir                                    ' Get next entry.
            Loop
            intDirNum = intDirNum + 1
        Loop
        intNum = intNum + 1
    Loop
    rsSongs.Save App.Path & "\Songs.lst", adPersistADTG     'save file to HD
    If MsgBox("Database list has been made.", vbOKOnly, "Finished") = vbOK And blnEmpty = True Then
        MsgBox "The text boxes left blank were changed to C:\", , "Empty Text Box"
    End If
    
    'save settings
    With rsSettings
    .AddNew
    intNum = 0
    .MoveFirst
    Do Until intNum = mintTotal             'retrieve field names from lbls
        .Fields(intNum) = lblName(intNum).Caption
        intNum = intNum + 1
    Loop
    Do Until intNum = 10                    'clear remaining fields
        .Fields(intNum) = ""
        intNum = intNum + 1
    Loop
    intNum = 10
    Do Until intNum = mintTotal + 10        'retrieve dir names from txts
        .Fields(intNum) = txtName(intNum - 10)
        intNum = intNum + 1
    Loop
    Do Until intNum = 20                    'clear remaining fields
        .Fields(intNum) = ""
        intNum = intNum + 1
    Loop
    .Fields(20) = mintTotal
    .Fields(21) = "Programmed By Drummer"        'Gloat section
    .Update
    .Save App.Path & "\Settings.ini", adPersistADTG   'save settings
    End With
    
    lblFull.Visible = True
    frmDir.Hide
    frmMusic.Show
    MousePointer = vbDefault
End Sub

'start over
Private Sub cmdRestart_Click()
    Dim intNum As Integer
    intNum = 1
    txtName(0).Text = ""
    txtName(0).SetFocus
    'set objects to visible or invisible
    Do Until intNum = 10
        txtName(intNum).Visible = False
        intNum = intNum + 1
    Loop
    intNum = 0
    Do Until intNum = 10
        lblName(intNum).Visible = False
        intNum = intNum + 1
    Loop
    lbl1.Visible = True
    lbl2.Visible = False
    lbl3.Visible = False
    cmdOK2.Visible = False
    cmdBrowse.Visible = False
    chkSubDir.Visible = False
    chkListDir.Visible = False
    cmdRestart2.Visible = False
    blnBoo = False
    Form_Activate
End Sub

'start over
Private Sub cmdRestart2_Click()
    cmdRestart_Click
End Sub

Private Sub Form_Activate()
    If cmdBrowse.Visible = True Then Exit Sub       'when browse window hides, it makes this form activate
    strDate = Date                                  'used "If" to skip Form_Activate so stuff stays the same
    blnBoo = False
    blnSettings = False
    Dim intNum As Integer
    Dim strTemp As String
    strTemp = Dir(App.Path & "\Songs.lst")
    Dim strSettings As String
    strSettings = Dir(App.Path & "\Settings.ini")
    mnuEdit.Visible = False
    
    'setup connections for ado controls
    Set rsSongs = New ADODB.Recordset
    Set rsSettings = New ADODB.Recordset
    
    'check for songs file
    If strTemp <> "" Then
        lblFull.Visible = True
        txtName(0).Visible = False
        cmdOK.Visible = False
        lbl1.Visible = False
        mnuEdit.Visible = True
    Else
        lblFull.Visible = False
        txtName(0).Visible = True
        cmdOK.Visible = True
        lbl1.Visible = True
        blnBoo = False
        cmdOK.Default = True
    End If
    
    'check for settings file
    If strSettings = "" Then
        blnSettings = True
        
        'create fields
        intNum = 0
        Do Until intNum = 22
            rsSettings.Fields.Append intNum, adVarChar, 100
            intNum = intNum + 1
        Loop
    End If
    rsSettings.Open App.Path & "\Settings.ini"
    If blnSettings = False Then
        txtName(0) = rsSettings.Fields(20)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadAll
End Sub

'delete songs.lst file
Private Sub mnuEditDelete_Click()
    Dim strDir As String
    strDir = App.Path & "\Songs.lst"
    If MsgBox("This will permanently remove the list from your hard drive." & vbCrLf _
        & "Do you want to continue?", vbOKCancel, "Sure?") = vbOK Then
        FSO.DeleteFile (strDir)
        Me.Hide
        frmDir.Show
    End If
End Sub

'show about
Private Sub mnuFileAbout_Click()
    MsgBox "Music List          v " & mstrVersion & vbCrLf & "By Drummer" & vbCrLf & vbCrLf & "Programmed: " & mintDate, vbOKOnly, "About Music List"
End Sub

'exit
Private Sub mnuFileExit_Click()
    UnloadAll
End Sub

'hide form
Private Sub mnuFileReturn_Click()
    frmDir.Hide
End Sub

'show form
Private Sub mnuWindowDir_Click()
    frmDir.Show vbModeless
End Sub

'show music list
Private Sub mnuWindowMusic_Click()
    frmMusic.Show vbModeless
End Sub

'select txtbox
Private Sub txtName_GotFocus(Index As Integer)
    'select text
    If txtName(Index) <> "" Then
        With txtName(Index)
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    strTextBoxNum = Index               'not used, don't remember why its here
    mintTXTNum = Index                  'to tell the browse form which txt to put dir in
End Sub
