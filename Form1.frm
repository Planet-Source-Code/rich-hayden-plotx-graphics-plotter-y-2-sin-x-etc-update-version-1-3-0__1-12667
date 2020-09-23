VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PlotX"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   618
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Controls:"
      Height          =   8655
      Left            =   8880
      TabIndex        =   68
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command4 
         Caption         =   "&ACCURACY/ SPEED"
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Height          =   615
         Left            =   320
         Picture         =   "Form1.frx":11C2
         ScaleHeight     =   555
         ScaleWidth      =   2235
         TabIndex        =   22
         Top             =   7920
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   315
         Index           =   7
         Left            =   2520
         TabIndex        =   21
         Top             =   7560
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Index           =   7
         Left            =   480
         TabIndex        =   20
         Top             =   7560
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   315
         Index           =   6
         Left            =   2520
         TabIndex        =   19
         Top             =   7200
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00FF8080&
         Height          =   285
         Index           =   6
         Left            =   480
         TabIndex        =   18
         Top             =   7200
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   315
         Index           =   5
         Left            =   2520
         TabIndex        =   17
         Top             =   6840
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H0000C0C0&
         Height          =   285
         Index           =   5
         Left            =   480
         TabIndex        =   16
         Top             =   6840
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   315
         Index           =   4
         Left            =   2520
         TabIndex        =   15
         Top             =   6480
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00C000C0&
         Height          =   285
         Index           =   4
         Left            =   480
         TabIndex        =   14
         Top             =   6480
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   13
         Top             =   6120
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   12
         Top             =   6120
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   315
         Index           =   2
         Left            =   2520
         TabIndex        =   11
         Top             =   5760
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H0000C000&
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   10
         Top             =   5760
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   9
         Top             =   5400
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   5400
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   7
         Top             =   5040
         Value           =   1  'Checked
         Width           =   255
      End
      Begin MSScriptControlCtl.ScriptControl ScriptControl1 
         Left            =   2160
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         AllowUI         =   -1  'True
         UseSafeSubset   =   -1  'True
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   5040
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Zoom In"
         Height          =   975
         Left            =   120
         Picture         =   "Form1.frx":58A4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Zoom Out"
         Height          =   975
         Left            =   1920
         Picture         =   "Form1.frx":6BA6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "DRAW &GRAPH(S)"
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Click on the 'x =' or 'y =' before the expression box to swap around the expression subject"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   83
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "y ="
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
         Index           =   7
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   80
         Top             =   7560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "y ="
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
         Index           =   6
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   79
         Top             =   7200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "y ="
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
         Index           =   5
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   78
         Top             =   6840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "y ="
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
         Index           =   4
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   77
         Top             =   6480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "y ="
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
         Index           =   3
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   76
         Top             =   6120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "y ="
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
         Index           =   2
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   75
         Top             =   5760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "y ="
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
         Index           =   1
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   74
         Top             =   5400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "y ="
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
         Index           =   0
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   73
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Expression(s):"
         Height          =   495
         Left            =   120
         TabIndex        =   72
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "To set the new scale, you must refresh the graphs!"
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   120
         TabIndex        =   71
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Scale:"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Co-ordinates: ----------------------"
         Height          =   615
         Left            =   120
         TabIndex        =   69
         Top             =   1920
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   8655
      Left            =   120
      ScaleHeight     =   8595
      ScaleWidth      =   8595
      TabIndex        =   2
      Top             =   120
      Width           =   8655
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rendering and Drawing Graph(s). Please wait...."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   840
         TabIndex        =   81
         Top             =   5400
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Label dumXn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   4080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label dumYn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4020
         TabIndex        =   66
         Top             =   8280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label dumYp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4380
         TabIndex        =   65
         Top             =   105
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblMxp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   10
         Left            =   7680
         TabIndex        =   64
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblMyn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   9
         Left            =   5520
         TabIndex        =   63
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lblMyn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   62
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblMyn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   61
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblMyn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   60
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblMyn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblMyn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   58
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblMyn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   57
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblMyn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   56
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblMyn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   55
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblMyn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   54
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblMyp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   53
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblMyp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblMyp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   51
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblMyp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   50
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblMyp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   49
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblMyp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   48
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblMyp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   47
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblMyp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   46
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblMyp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   45
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblMyp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   10
         Left            =   1680
         TabIndex        =   44
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblMxn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   9
         Left            =   2760
         TabIndex        =   43
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblMxn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   8
         Left            =   4320
         TabIndex        =   42
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblMxn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   41
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label lblMxn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   40
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label lblMxn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   39
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label lblMxn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   38
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label lblMxn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   37
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label lblMxn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   36
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label lblMxn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   35
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label lblMxn 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   10
         Left            =   4800
         TabIndex        =   34
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label lblMxp 
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   9
         Left            =   8160
         TabIndex        =   33
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblMxp 
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   8
         Left            =   7710
         TabIndex        =   32
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblMxp 
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   7
         Left            =   7260
         TabIndex        =   31
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblMxp 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   6810
         TabIndex        =   30
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblMxp 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   29
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblMxp 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   28
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblMxp 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   27
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblMxp 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   26
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblMxp 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   25
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label dumXp 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8640
         TabIndex        =   24
         Top             =   4320
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Copyright (c) 2000 Richard Hayden. All Rights Reserved."
      Height          =   615
      Left            =   8880
      TabIndex        =   82
      Top             =   8880
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Ready..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Top             =   8880
      Width           =   8895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xmax, ymax, xmin, ymin, Y, X, sm, pi, accuracy As Currency
Dim equation(1 To 10) As Currency
Dim c1, val, x2, x3 As Integer
Dim ystr As String
Dim blnError As Boolean

Dim curVal As String

Public Sub Update_Settings()
    On Error GoTo EH
    Frame1.Enabled = False
    
    If GetIniInfo(App.Path & "\plotx.ini", "GRAPH DRAWING", "Accuracy", "3") = "1" Then
        accuracy = 0.1
    ElseIf GetIniInfo(App.Path & "\plotx.ini", "GRAPH DRAWING", "Accuracy", "3") = "2" Then
        accuracy = 0.01
    ElseIf GetIniInfo(App.Path & "\plotx.ini", "GRAPH DRAWING", "Accuracy", "3") = "3" Then
        accuracy = 0.001
    ElseIf GetIniInfo(App.Path & "\plotx.ini", "GRAPH DRAWING", "Accuracy", "3") = "4" Then
        accuracy = 0.0001
    ElseIf GetIniInfo(App.Path & "\plotx.ini", "GRAPH DRAWING", "Accuracy", "3") = "5" Then
        accuracy = 0.000001
    End If
    
    Frame1.Enabled = True
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: O" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub Check1_Click(Index As Integer)
    On Error GoTo EH
    If Check1(Index).Value = 1 Then
        Text1(Index).Enabled = True
        Text1(Index).BackColor = &H80000005
        Label5(Index).Visible = True
    Else
        Text1(Index).Enabled = False
        Text1(Index).BackColor = &H8000000F
        Label5(Index).Visible = False
    End If
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: A" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub Combo1_Click()
    On Error GoTo EH
    If Not (IsNumeric(Combo1.Text)) Then
        Combo1.Text = curVal
        Picture1.SetFocus
    ElseIf CInt(Combo1.Text) > 100 Or CInt(Combo1.Text) < 1 Then
        Combo1.Text = curVal
        Picture1.SetFocus
    End If
    Combo1.Text = Abs(CInt(Combo1.Text))
    curVal = Combo1.Text
    If Not (CInt(Combo1.Text) = val) Then
        Label3.Visible = True
    End If
    val = CInt(Combo1.Text)
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: B" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub Combo1_Change()
    On Error GoTo EH
    If Not (IsNumeric(Combo1.Text)) Then
        Combo1.Text = curVal
        Picture1.SetFocus
    ElseIf CInt(Combo1.Text) > 100 Or CInt(Combo1.Text) < 1 Then
        Combo1.Text = curVal
        Picture1.SetFocus
    End If
    Combo1.Text = Abs(CInt(Combo1.Text))
    curVal = Combo1.Text
    If Not (CInt(Combo1.Text) = val) Then
        Label3.Visible = True
    End If
    val = CInt(Combo1.Text)
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: C" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub Command1_Click()
    blnError = False
    Frame1.Enabled = False
    Label7.Visible = True
    Dim c, temp, m As Currency
    On Error Resume Next
    Label3.Visible = False
    xmax = CCur(Combo1.Text)
    xmin = -CCur(Combo1.Text)
    ymax = CCur(Combo1.Text)
    ymin = -CCur(Combo1.Text)
    X = 0
    Y = 0
    Picture1.Cls
    c = 0
    temp = 0
    sm = 10
    Picture1.ForeColor = &HFF&
    Picture1.Scale (xmin, ymax)-(xmax, ymin)
    Picture1.Line (0, ymin)-(0, ymax)
    Picture1.Line (xmin, 0)-(xmax, 0)
    Picture1.ForeColor = &HFF0000
    
    m = 0
    If xmax < 0 Then
        temp = (-xmax) / sm
    Else
        temp = xmax / sm
    End If
    If xmax < 0 Then
        c = -xmax
    Else
        c = xmax
    End If
    While c >= xmin
        m = m + 1
        If ymax < 0 Then
            Picture1.Line (c, (-ymax) / 100)-(c, -((-ymax) / 100))
        Else
            Picture1.Line (c, ymax / 100)-(c, -(ymax / 100))
        End If
        c = c - temp
    Wend
    
    If ymax < 0 Then
        temp = (-ymax) / sm
    Else
        temp = ymax / sm
    End If
    If ymax < 0 Then
        c = -ymax
    Else
        c = ymax
    End If
    While c >= ymin
        If xmax < 0 Then
            Picture1.Line ((-xmax) / 100, c)-(-((-xmax) / 100), c)
        Else
            Picture1.Line (xmax / 100, c)-(-(xmax / 100), c)
        End If
        c = c - temp
    Wend
    Picture1.ForeColor = &H80&
    
    Dim l, o As Integer
    l = 11
    While l > 1
        l = l - 1
        lblMxp(l).Caption = (xmax / 10) * l
        lblMxp(l).ToolTipText = lblMxp(l).Caption
    Wend
    l = 0
    o = 0
    While l > -10
        l = l - 1
        o = o + 1
        lblMxn(o).Caption = (xmax / 10) * l
        lblMxn(o).ToolTipText = lblMxn(o).Caption
    Wend
    l = 11
    While l > 1
        l = l - 1
        lblMyp(l).Caption = (xmax / 10) * l
        lblMyp(l).ToolTipText = lblMyp(l).Caption
    Wend
    l = 0
    o = 0
    While l > -10
        l = l - 1
        o = o + 1
        lblMyn(o).Caption = (xmax / 10) * l
        lblMyn(o).ToolTipText = lblMyn(o).Caption
    Wend
    
    Picture1.ForeColor = &H80000012
    x2 = 0
    While x2 <= 7 And Not (blnError)
    If Check1(x2).Value = 1 Then
    
    If Label5(x2).Caption = "y =" Then
    
    X = xmax
    Label1.Caption = "Rendering and Drawing Graph (Expression: " & x2 + 1 & ")..."
    DoEvents
    
    ScriptControl1.Reset
    ScriptControl1.Language = "VBScript"
    ScriptControl1.ExecuteStatement ("Dim x")
    ScriptControl1.ExecuteStatement ("x = " & xmax)
    While X >= xmin And Not (blnError)
        Y = ScriptControl1.Eval(Text1(x2).Text)
        If Not (Y > ymax Or Y < ymin) Then
            'SetPixel Picture1.hdc, x, y, &H80000007
            Picture1.PSet (X, Y), Text1(x2).ForeColor
        End If
        ScriptControl1.ExecuteStatement ("x = x - " & accuracy)
        X = X - accuracy
    Wend
    
    Else
    
    Y = ymax
    Label1.Caption = "Rendering and Drawing Graph (Expression: " & x2 + 1 & ")..."
    DoEvents
    
    ScriptControl1.Reset
    ScriptControl1.Language = "VBScript"
    ScriptControl1.ExecuteStatement ("Dim y")
    ScriptControl1.ExecuteStatement ("y = " & ymax)
    While Y >= ymin And Not (blnError)
        X = ScriptControl1.Eval(Text1(x2).Text)
        If Not (X > xmax Or X < xmin) Then
            'SetPixel Picture1.hdc, x, y, &H80000007
            Picture1.PSet (X, Y), Text1(x2).ForeColor
        End If
        ScriptControl1.ExecuteStatement ("y = y - " & accuracy)
        Y = Y - accuracy
    Wend
    
    End If
    
    Label1.Caption = "Graph Complete, moving onto next..."
    End If
    x2 = x2 + 1
    Wend
    If blnError Then
        blnError = False
        Label1.Caption = "Graph(s) Incomplete, due to an error in expression " & x3 & ", READY..."
        Label7.Visible = False
        Frame1.Enabled = True
    Else
        Label1.Caption = "Graph(s) Complete, READY..."
        Label7.Visible = False
        Frame1.Enabled = True
    End If
End Sub

Private Sub Command2_Click()
    On Error GoTo EH
    Combo1.Text = CInt(Combo1.Text) - 1
    Call Command1_Click
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: D" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub Command3_Click()
    On Error GoTo EH
    Combo1.Text = CInt(Combo1.Text) + 1
    Call Command1_Click
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: E" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub Command4_Click()
    Call frmOptions.Form_Load
    frmOptions.Show 1
End Sub

Private Sub Form_Load()
    Dim c, temp, m As Currency
    curVal = "10"
    val = 10
    c1 = 0
    pi = 4 * Atn(1)
    While c1 < 100
        c1 = c1 + 1
        Combo1.AddItem c1
    Wend
    Combo1.Text = "10"
    On Error Resume Next
    xmax = CCur(Combo1.Text)
    xmin = -CCur(Combo1.Text)
    ymax = CCur(Combo1.Text)
    ymin = -CCur(Combo1.Text)
    Call Update_Settings
    Picture1.Cls
    c = 0
    temp = 0
    X = 0
    sm = 10
    Picture1.ForeColor = &HFF&
    Picture1.Scale (xmin, ymax)-(xmax, ymin)
    Picture1.Line (0, ymin)-(0, ymax)
    Picture1.Line (xmin, 0)-(xmax, 0)
    Picture1.ForeColor = &HFF0000
    
    m = 0
    If xmax < 0 Then
        temp = (-xmax) / sm
    Else
        temp = xmax / sm
    End If
    If xmax < 0 Then
        c = -xmax
    Else
        c = xmax
    End If
    While c >= xmin
        m = m + 1
        If ymax < 0 Then
            Picture1.Line (c, (-ymax) / 100)-(c, -((-ymax) / 100))
        Else
            Picture1.Line (c, ymax / 100)-(c, -(ymax / 100))
        End If
        c = c - temp
    Wend
    
    m = 0
    If ymax < 0 Then
        temp = (-ymax) / sm
    Else
        temp = ymax / sm
    End If
    If ymax < 0 Then
        c = -ymax
    Else
        c = ymax
    End If
    While c >= ymin
        If xmax < 0 Then
            Picture1.Line ((-xmax) / 100, c)-(-((-xmax) / 100), c)
        Else
            Picture1.Line (xmax / 100, c)-(-(xmax / 100), c)
        End If
        c = c - temp
    Wend
    
    Dim l, o As Integer
    l = 11
    While l > 1
        l = l - 1
        lblMxp(l).Left = l
        lblMxp(l).Top = dumXp.Top
        lblMxp(l).Caption = l
        lblMxp(l).ToolTipText = lblMxp(l).Caption
    Wend
    l = 0
    o = 0
    While l > -10
        l = l - 1
        o = o + 1
        lblMxn(o).Left = l
        lblMxn(o).Top = dumXn.Top
        lblMxn(o).Caption = l
        lblMxn(o).ToolTipText = lblMxn(o).Caption
    Wend
    l = 11
    While l > 1
        l = l - 1
        lblMyp(l).Top = l
        lblMyp(l).Left = dumYp.Left
        lblMyp(l).Caption = l
        lblMyp(l).ToolTipText = lblMyp(l).Caption
    Wend
    l = 0
    o = 0
    While l > -10
        l = l - 1
        o = o + 1
        lblMyn(o).Top = l
        lblMyn(o).Left = dumYn.Left
        lblMyn(o).Caption = l
        lblMyn(o).ToolTipText = lblMyn(o).Caption
    Wend
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    Label2.Caption = "Co-ordinates: ----------------------"
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: F" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    Label2.Caption = "Co-ordinates: ----------------------"
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: G" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    Label2.Caption = "Co-ordinates: ----------------------"
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: H" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub Label5_Click(Index As Integer)
    If Label5(Index).Caption = "y =" Then
        Label5(Index).Caption = "x ="
    Else
        Label5(Index).Caption = "y ="
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo EH
    Label2.Caption = "Co-ordinates: (" & CCur(X) & ", " & CCur(Y) & ")"
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: I" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub Picture2_Click()
    On Error GoTo EH
    frmAbout.Show 1
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: J" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Private Sub ScriptControl1_Error()
    blnError = True
    x3 = x2 + 1
    MsgBox "There was an error in expression " & x3 & ". The graph of expression " & x3 & " may not be complete or may be incorrect. Any proceeding expression's graphs will not have been drawn.", vbApplicationModal + vbInformation + vbOKOnly, "PlotX"
    Label1.Caption = "Graph(s) Incomplete, due to an error in expression " & x3 & ", READY..."
    Label7.Visible = False
    Frame1.Enabled = True
End Sub
