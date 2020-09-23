VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstSearch 
      Height          =   4740
      ItemData        =   "frmSearch.frx":0000
      Left            =   0
      List            =   "frmSearch.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   8775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Locations"
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   8775
      Begin VB.OptionButton optCheckAll 
         Caption         =   "Uncheck All"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optCheckAll 
         Caption         =   "Check All"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox chkDirCheck 
         Caption         =   "Check10"
         Height          =   255
         Index           =   9
         Left            =   6600
         TabIndex        =   12
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkDirCheck 
         Caption         =   "Check9"
         Height          =   255
         Index           =   8
         Left            =   6600
         TabIndex        =   11
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkDirCheck 
         Caption         =   "Check8"
         Height          =   255
         Index           =   7
         Left            =   6600
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkDirCheck 
         Caption         =   "Check7"
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkDirCheck 
         Caption         =   "Check6"
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkDirCheck 
         Caption         =   "Check5"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkDirCheck 
         Caption         =   "Check4"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   6
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkDirCheck 
         Caption         =   "Check3"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   5
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkDirCheck 
         Caption         =   "Check2"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkDirCheck 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblListCount 
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rsSongs As ADODB.Recordset
Attribute rsSongs.VB_VarHelpID = -1
Dim WithEvents rsSettings As ADODB.Recordset
Attribute rsSettings.VB_VarHelpID = -1

Private Sub chkDirCheck_Click(Index As Integer)
    txtSearch.SetFocus
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo HandleError
    Dim intNum, intNum2 As Integer
    Dim strNum, strSearch As String
    intNum = 0
    intNum2 = 0
    MousePointer = vbHourglass
    strSearch = txtSearch
    rsSettings.MoveFirst
    lstSearch.Clear
    If txtSearch = "" Then
        MsgBox "Enter a string into the text box", , "Empty String"
        MousePointer = vbDefault
        Exit Sub
    End If
    
    Do Until intNum = 10
        rsSongs.MoveFirst
        If chkDirCheck(intNum).Value = Checked Then
            Do Until rsSongs.EOF
                strNum = intNum
                strSearch = "[" & intNum & "] Like '*" & txtSearch.Text & "*'"
                rsSongs.Find strSearch
                If rsSongs.EOF Then
                Else
                If mstrOpen = "Songs.lst" Then
                    lstSearch.AddItem rsSettings.Fields(intNum + 10) & rsSongs.Fields(strNum)
                Else
                    lstSearch.AddItem rsSongs.Fields(strNum)
                End If
                rsSongs.MoveNext
                End If
            Loop
        End If
        intNum = intNum + 1
    Loop
    lblListCount.Caption = lstSearch.ListCount & " Files in List"
    MousePointer = vbDefault
    txtSearch.SetFocus
    Exit Sub
HandleError:
    MsgBox Err.Description
    MousePointer = vbDefault
    txtSearch.SetFocus
End Sub

Private Sub Form_Activate()
    Set rsSongs = New ADODB.Recordset
    Set rsSettings = New ADODB.Recordset
    Dim intNum As Integer
    Dim strNum As String
    intNum = 0
    
    Dim strTemp As String
    Dim strTemp2 As String
    strTemp = Dir(App.Path & "\" & mstrOpen)
    strTemp2 = Dir(App.Path & "\Settings.ini")
    If strTemp2 <> "" Then
        rsSettings.Open App.Path & "\Settings.ini"
    Else
        MsgBox "Settings file is missing.", , "Missing File"
    End If
    If strTemp <> "" Then
        rsSongs.Open App.Path & "\" & mstrOpen
    Else
        MsgBox "Music list file is missing.", , "Missing File"
    End If
    
    rsSongs.MoveFirst
    Do Until intNum = 10
        chkDirCheck(intNum).Visible = False
        intNum = intNum + 1
    Loop
    intNum = 0
    Do Until intNum = 10
        strNum = intNum
        If rsSongs.Fields(strNum) <> "" And rsSongs.Fields(strNum) <> " " Then
            chkDirCheck(intNum).Caption = rsSongs.Fields(strNum)
            chkDirCheck(intNum).Visible = True
        End If
        intNum = intNum + 1
    Loop
End Sub

Private Sub optCheckAll_Click(Index As Integer)
    Dim intNum As Integer
    intNum = 0
    Do Until intNum = 10
        If Index = 0 Then
            chkDirCheck(intNum) = Checked
        ElseIf Index = 1 Then
            chkDirCheck(intNum) = Unchecked
        End If
        intNum = intNum + 1
    Loop
    txtSearch.SetFocus
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch_Click
    End If
End Sub
