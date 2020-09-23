VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox lstFile 
      Height          =   7470
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton cmdFun 
      Caption         =   "&Fun"
      Height          =   255
      Left            =   7800
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboDir 
      Height          =   315
      ItemData        =   "frmBrowse.frx":0000
      Left            =   0
      List            =   "frmBrowse.frx":0002
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "&Up"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtDir 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8295
   End
   Begin VB.ListBox lstDir 
      Height          =   7470
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblDir 
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblFile 
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim MyPath, MyName, strAddress As String

Private Sub cboDir_Click()
    strAddress = cboDir.Text
    txtDir = cboDir.Text
    ListDir
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFun_Click()
    Dim x As Integer
    x = 99
    Do While x > 0
    MsgBox x & " bottles of beer on the wall, " & x & " bottles of beer. Take on down, pass it around, " & x - 1 & " bottles of beer on the wall."
    x = x - 1
    Loop
End Sub

Private Sub cmdOK_Click()
    Me.Hide
    frmDir.txtName(mintTXTNum) = txtDir
End Sub

Private Sub cmdOpen_Click()
    cmdOpen.Enabled = False
    ListDir
End Sub

Private Sub cmdUp_Click()
    If Len(strAddress) < 4 Then
        'MsgBox "Cannot continue in this direction.", , "End Of The Line"
    Else
    Dim intNum, intStr As Integer
    intNum = 0
    intStr = Len(strAddress) - 1
    strAddress = Left(strAddress, intStr)
    Do Until Right(strAddress, 1) = "\"
        intStr = Len(strAddress) - 1
        strAddress = Left(strAddress, intStr)
    Loop
    ListDir
    End If
End Sub

Private Sub Form_Load()
    strAddress = txtDir.Text = "C:\"
    cmdOpen.Enabled = False
    cboDir.AddItem "A:\"
    cboDir.AddItem "B:\"
    cboDir.AddItem "C:\"
    cboDir.AddItem "D:\"
    cboDir.AddItem "E:\"
    cboDir.AddItem "F:\"
    cboDir.AddItem "G:\"
    cboDir.AddItem "H:\"
    cboDir.AddItem "I:\"
    cboDir.AddItem "J:\"
    cboDir.AddItem "K:\"
    cboDir.AddItem "L:\"
    cboDir.AddItem "M:\"
    cboDir.AddItem "N:\"
    cboDir.AddItem "O:\"
    cboDir.AddItem "P:\"
    cboDir.AddItem "Q:\"
    cboDir.AddItem "R:\"
    cboDir.AddItem "S:\"
    cboDir.AddItem "T:\"
    cboDir.AddItem "U:\"
    cboDir.AddItem "V:\"
    cboDir.AddItem "W:\"
    cboDir.AddItem "X:\"
    cboDir.AddItem "Y:\"
    cboDir.AddItem "Z:\"
    cboDir.ListIndex = 2
    ListDir
End Sub

Private Sub lstFile_DblClick()
    ShellExecute 0, "open", txtDir & "\" & lstFile.Text, "", txtDir, 1
End Sub

Private Sub txtDir_LostFocus()
    strAddress = txtDir
    ListDir
End Sub

Private Sub lstDir_Click()
    Dim strTemp As String
    strTemp = txtDir
    If Right(strTemp, 1) = "\" Then
        strAddress = txtDir & lstDir.Text
    Else
        strAddress = txtDir & "\" & lstDir.Text
    End If
    cmdOpen.Enabled = True
End Sub

Private Sub lstDir_DblClick()
    cmdOpen.Enabled = False
    Dim strTemp As String
    strTemp = txtDir
    If Right(strTemp, 1) = "\" Then
        strAddress = txtDir & lstDir.Text
    Else
        strAddress = txtDir & "\" & lstDir.Text
    End If
    ListDir
End Sub

Private Sub lstDir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ListDir
    End If
End Sub

Private Sub ListDir()
    On Error GoTo HandleError
    MousePointer = vbHourglass
    Dim strDir() As String  'undefined number of subdirs
    Dim strName() As String 'undefined number of subdir names
    Dim intDirCount As Integer   'number of dir
    Dim intNum As Integer   'loop increment dir number
    intDirCount = 1         'must set before ReDimn'
    ReDim strDir(intDirCount)    'set undefined string to a number
    ReDim strName(intDirCount)   'set undefined string to a number
    lstDir.Clear
    lstFile.Clear
    strDir(0) = strAddress                   ' Set the path.
    intNum = 0
    intDirCount = 1                            'must set to 1 before each sub search, in case following loop is placed within loop
    Do Until intNum = intDirCount        'loop till all directories are covered
        If Right(strDir(intNum), 1) = "\" Then
            MyPath = strDir(intNum) & strName(intNum)
        Else
            MyPath = strDir(intNum) & strName(intNum) & "\" 'concatinate dir from temp string
        End If
        MyName = Dir(MyPath, vbDirectory)               'generate name from path
        Do While MyName <> ""                           'loop if it exists
            If MyName <> "." And MyName <> ".." Then    'if its a file
                If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then    'check if its dir
                    strDir(intDirCount) = MyPath         'add path to temp dir
                    strName(intDirCount) = MyName        'add file to temp dir
                    'If chkSubDir = Checked Then
                        'intDirCount = intDirCount + 1         'increment if you want subdir
                    'End If
                    ReDim Preserve strDir(intDirCount)   'set undefined string to a number
                    ReDim Preserve strName(intDirCount)  'set undefined string to a number
                    
                    'If optFiles = False Then
                        lstDir.AddItem MyName        'add to the list if its a dir
                    'End If
                Else
                    'If optSub = False Then
                        lstFile.AddItem MyName       'add to the list if its a file
                    'End If
                End If
            End If
            MyName = Dir   ' Get next entry.
        Loop
        intNum = intNum + 1     'increment
    Loop
    txtDir = strAddress
    lblDir.Caption = lstDir.ListCount & " Folders" 'show count
    lblFile.Caption = lstFile.ListCount & " Files" 'show count
    MousePointer = vbDefault
    Exit Sub
HandleError:
    MsgBox "Drive is not ready.", , "Try Later"
    MousePointer = vbDefault
End Sub
