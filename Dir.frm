VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.ListBox lstDir 
      Height          =   5325
      ItemData        =   "Dir.frx":0000
      Left            =   120
      List            =   "Dir.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Click"
      Default         =   -1  'True
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblList 
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Number of Files:"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Directory:"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyFile, MyPath, MyName

Private Sub cmdRun_Click()
    Clear
    ' Display the names in C:\ that represent directories.
    MyPath = txtDir.Text                    ' Set the path.
    MyName = Dir(MyPath)                    ' Retrieve the first entry.
    Do While MyName <> ""                   ' Start the loop.
        ' Ignore the current directory and the encompassing directory.
        If MyName <> "." And MyName <> ".." Then
            lstDir.AddItem MyName
        End If
        MyName = Dir   ' Get next entry.
    Loop
    lblList.Caption = lstDir.ListCount
End Sub

Private Sub Form_Load()
    txtDir.Text = "C:\Backup\Installation\"
End Sub

Private Sub Clear()
    'clear list
    
    lstDir.Clear
End Sub
