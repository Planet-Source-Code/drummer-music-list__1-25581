VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Select a file to open."
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ListBox lstOpen 
         Height          =   3180
         ItemData        =   "frmOpen.frx":0442
         Left            =   120
         List            =   "frmOpen.frx":0444
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFile, strDirectory As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    If lstOpen.Text = "" Then
        mstrOpen = "Songs.lst"
    Else
        mstrOpen = lstOpen.Text
    End If
    Unload Me
End Sub

Private Sub Form_Load()             'for recursing directories
    strDirectory = App.Path & "\"      ' Set the path.
    strFile = Dir(strDirectory)   ' Retrieve the first entry.
    Do While strFile <> ""               ' Start the loop.
        ' Ignore the current directory and the encompassing directory.
        If strFile <> "." And strFile <> ".." Then
            If Right(strFile, 4) = ".lst" Then
                lstOpen.AddItem strFile
            End If
        End If
        strFile = Dir                                    ' Get next entry.
    Loop
    
End Sub

Private Sub lstOpen_DblClick()
    cmdOpen_Click
End Sub
