VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PlotX - Speed/Accuracy"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   3720
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Speed/Accuracy"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "I&ncredibly Slow/Maximum Accuracy"
         Height          =   615
         Index           =   5
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Very Slow/Very Accurate"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Slow/Accurate"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   1
         Top             =   1680
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Quite Fast/Quite Innaccurate"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Very Fast/Very &Inaccurate"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   2880
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Command2_Click()
    On Error GoTo EH
    Dim strTemp As String
    Dim intTemp As Integer
    Frame1.Enabled = False
    
    Do While intTemp < 5
        intTemp = intTemp + 1
        If Option1(intTemp).Value = True Then
            WriteIniInfo "GRAPH DRAWING", "Accuracy", CStr(intTemp), App.Path & "\plotx.ini"
            Exit Do
        End If
    Loop
    
    Frame1.Enabled = True
    Call Form1.Update_Settings
    Me.Hide
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: N" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Public Sub Form_Load()
    On Error GoTo EH
    Dim strTemp As String
    Dim intTemp As Integer
    
    strTemp = GetIniInfo(App.Path & "\plotx.ini", "GRAPH DRAWING", "Accuracy", "3")
    intTemp = CInt(strTemp)
    Option1(intTemp).Value = True
    
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: K" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub
