VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Special Relativity"
   ClientHeight    =   6930
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7620
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Other"
      Height          =   2895
      Left            =   6240
      TabIndex        =   63
      Top             =   2880
      Width           =   1335
      Begin VB.CommandButton Command5 
         Caption         =   "Other Help"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   2040
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   480
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   ".lst"
         DialogTitle     =   "Open Special Relativity List"
         FileName        =   "SpecRel1.lst"
         Filter          =   "SR List Files | *.lst"
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Explanations"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Other Relativity Problems"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   66
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reset"
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ListBox Status 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "Form2.frx":058A
      Left            =   0
      List            =   "Form2.frx":058C
      TabIndex        =   1
      Top             =   5880
      Width           =   7575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Calculations"
      Height          =   2895
      Left            =   0
      TabIndex        =   30
      Top             =   2880
      Width           =   6135
      Begin VB.OptionButton Option13 
         Caption         =   "Find t2' time dil"
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Find t2 time dil"
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Find x2' len contr"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Find x2 len contr"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Find v Time Dilation"
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Find v Len Contraction"
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   240
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Results:"
         Height          =   1335
         Left            =   2880
         TabIndex        =   60
         Top             =   1440
         Width           =   3015
         Begin VB.CheckBox assign 
            Caption         =   "Assign Value"
            Height          =   375
            Left            =   1560
            TabIndex        =   65
            Top             =   840
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Answer 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   62
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.ListBox Missing 
         Height          =   1035
         ItemData        =   "Form2.frx":058E
         Left            =   1440
         List            =   "Form2.frx":0590
         TabIndex        =   58
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ListBox Uses 
         Height          =   1035
         ItemData        =   "Form2.frx":0592
         Left            =   120
         List            =   "Form2.frx":0594
         TabIndex        =   56
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Find t2'"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Find t2"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Find x2'"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Find x2"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Find ux'"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Find ux"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Find v (vel trans.)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "You're missing:"
         Height          =   255
         Left            =   1440
         TabIndex        =   59
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "This uses:"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Values (Position, Velocity, Time)"
      Height          =   2895
      Left            =   4320
      TabIndex        =   2
      Top             =   0
      Width           =   3255
      Begin VB.ListBox lstSave 
         Height          =   1425
         ItemData        =   "Form2.frx":0596
         Left            =   960
         List            =   "Form2.frx":059D
         TabIndex        =   68
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Set"
         Height          =   255
         Left            =   2040
         TabIndex        =   48
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Set"
         Height          =   255
         Left            =   2040
         TabIndex        =   47
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox TXTT2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox TXTT2PRIME 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Set"
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Set"
         Height          =   255
         Left            =   2040
         TabIndex        =   39
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Set"
         Height          =   255
         Left            =   2040
         TabIndex        =   38
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Set"
         Height          =   255
         Left            =   2040
         TabIndex        =   37
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Set"
         Height          =   255
         Left            =   2040
         TabIndex        =   36
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TXTUX 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox TXTX2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox TXTX2PRIME 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox TXTUXPRIME 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox TXTV 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "t2' = "
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label26 
         Caption         =   "t2 = "
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "time - units"
         Height          =   255
         Left            =   960
         TabIndex        =   44
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "time - units"
         Height          =   255
         Left            =   960
         TabIndex        =   43
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "c"
         Height          =   255
         Left            =   960
         TabIndex        =   29
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "ux' = "
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "c"
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "c"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "light - units"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "light - units"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "x2 = "
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "x2' = "
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "ux = "
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "v = "
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Illustration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Label t2prime 
         BackStyle       =   0  'Transparent
         Caption         =   "{t2'}"
         Height          =   255
         Left            =   2040
         TabIndex        =   35
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "{t1'} = 0"
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "{t1} = 0"
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label t2 
         BackStyle       =   0  'Transparent
         Caption         =   "{t2}"
         Height          =   255
         Left            =   1920
         TabIndex        =   32
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label ux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "{ux}"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "{x1'} = 0"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label x2prime 
         BackStyle       =   0  'Transparent
         Caption         =   "{x2'}"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "{x}1 = 0"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   2040
         X2              =   2040
         Y1              =   2280
         Y2              =   240
      End
      Begin VB.Label x2 
         BackStyle       =   0  'Transparent
         Caption         =   "{x2}"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Shape Shape4 
         Height          =   2535
         Left            =   120
         Top             =   240
         Width           =   3975
      End
      Begin VB.Line Line11 
         X1              =   960
         X2              =   960
         Y1              =   240
         Y2              =   2760
      End
      Begin VB.Line Line12 
         X1              =   120
         X2              =   4080
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   135
      End
      Begin VB.Shape Shape5 
         Height          =   1455
         Left            =   240
         Shape           =   2  'Oval
         Top             =   480
         Width           =   3135
      End
      Begin VB.Line Line13 
         X1              =   960
         X2              =   960
         Y1              =   600
         Y2              =   1800
      End
      Begin VB.Line Line14 
         X1              =   360
         X2              =   3240
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "y'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CS'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   2280
         Width           =   255
      End
      Begin VB.Shape Shape6 
         Height          =   255
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   840
         Width           =   255
      End
      Begin VB.Line Line15 
         X1              =   3360
         X2              =   3960
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line16 
         X1              =   3840
         X2              =   3975
         Y1              =   1335
         Y2              =   1200
      End
      Begin VB.Line Line17 
         X1              =   3840
         X2              =   3975
         Y1              =   1080
         Y2              =   1200
      End
      Begin VB.Label v 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.Line Line18 
         X1              =   1440
         X2              =   1920
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line19 
         X1              =   1800
         X2              =   1935
         Y1              =   1095
         Y2              =   960
      End
      Begin VB.Line Line20 
         X1              =   1800
         X2              =   1935
         Y1              =   840
         Y2              =   960
      End
      Begin VB.Label uxprime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "{ux'}"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuHelpLenContr 
         Caption         =   "Derivation of &Length Contraction"
      End
      Begin VB.Menu mnuHelpTimeDil 
         Caption         =   "Derivation of &Time Dilation"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command10_Click()
Call ux_Click
End Sub

Private Sub Command11_Click()
Call v_Click
End Sub

Private Sub Command12_Click()
Call t2prime_Click
End Sub

Private Sub Command13_Click()
Call t2_Click
End Sub

Private Sub Command2_Click()
Dim myVar As Variant
myVar = MsgBox("Are you sure you want to reset? Click Yes to reset, No to save and then reset, or cancel to exit!", vbYesNoCancel)
If myVar = vbYes Then
m:
'Command3.Enabled = False
v.Caption = "v"
TXTV.Text = ""

x2.Caption = "{x2}"
TXTX2.Text = ""
x2prime.Caption = "{x2'}"
TXTX2PRIME.Text = ""

t2.Caption = "{t2}"
TXTT2.Text = ""
t2prime.Caption = "{t2'}"
TXTT2PRIME.Text = ""

ux.Caption = "{ux}"
TXTUX.Text = ""
uxprime.Caption = "{ux'}"
TXTUXPRIME.Text = ""

Call Option1_Click
Option1.Value = True

Exit Sub

ElseIf myVar = vbNo Then
Call mnuFileSave_Click
GoTo m:

ElseIf myVar = vbCancel Then
Exit Sub
End If

End Sub

Private Sub Command3_Click()
If Not Missing.ListCount = 0 Then
    MsgBox "You are still missing values!"
    Exit Sub
End If

If Option1.Value = True Then
Answer.Caption = Round(CalcVRel(TXTUX.Text, TXTUXPRIME.Text), 3) & "c"
    If assign.Value = 1 Then
        TXTV.Text = Left(Answer.Caption, Len(Answer.Caption) - 1)
        v.Caption = "v = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [V SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option2.Value = True Then
Answer.Caption = Round(CalcVLenCont(TXTX2.Text, TXTX2PRIME.Text), 3) & "c"
    If assign.Value = 1 Then
        TXTV.Text = Left(Answer.Caption, Len(Answer.Caption) - 1)
        v.Caption = "v = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [V SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option3.Value = True Then
Answer.Caption = Round(CalcUX(TXTUXPRIME.Text, TXTV.Text), 3) & "c"
    If assign.Value = 1 Then
        TXTUX.Text = Left(Answer.Caption, Len(Answer.Caption) - 1)
        ux.Caption = "{ux} = " + Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [UX SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option4.Value = True Then
Answer.Caption = Round(CalcUXPrime(TXTUX.Text, TXTV.Text), 3) & "c"
    If assign.Value = 1 Then
        TXTUXPRIME.Text = Left(Answer.Caption, Len(Answer.Caption) - 1)
        uxprime.Caption = "{ux'} = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [UX' SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option5.Value = True Then
Answer.Caption = Round(CalcX(TXTX2PRIME.Text, TXTV.Text, TXTT2PRIME.Text), 3) & " lu"
    If assign.Value = 1 Then
        TXTX2.Text = Left(Answer.Caption, Len(Answer.Caption) - 3)
        x2.Caption = "{x2} = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [X2 SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option6.Value = True Then
Answer.Caption = Round(CalcXPrime(TXTX2.Text, TXTV.Text, TXTT2.Text), 3) & " lu"
    If assign.Value = 1 Then
        TXTX2PRIME.Text = Left(Answer.Caption, Len(Answer.Caption) - 3)
        x2prime.Caption = "{x2'} = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [X2' SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option7.Value = True Then
Answer.Caption = Round(CalcT(TXTT2PRIME.Text, TXTV.Text, TXTX2PRIME.Text), 3) & " tu"
    If assign.Value = 1 Then
        TXTT2.Text = Left(Answer.Caption, Len(Answer.Caption) - 3)
        t2.Caption = "{t2} = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [T2 SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option8.Value = True Then
Answer.Caption = Round(CalcTPrime(TXTT2.Text, TXTV.Text, TXTX2.Text), 3) & " lu"
    If assign.Value = 1 Then
        TXTT2PRIME.Text = Left(Answer.Caption, Len(Answer.Caption) - 3)
        t2prime.Caption = "{t2'} = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [T2' SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option9.Value = True Then
Answer.Caption = Round(CalcVTimeDil(TXTT2.Text, TXTT2PRIME.Text), 3) & "c"
    If assign.Value = 1 Then
        TXTV.Text = Left(Answer.Caption, Len(Answer.Caption) - 1)
        v.Caption = "v = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [V SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option10.Value = True Then
    Answer.Caption = Round(CalcL(TXTX2PRIME.Text, TXTV.Text), 3) & " lu"
    If assign.Value = 1 Then
        TXTX2.Text = Left(Answer.Caption, Len(Answer.Caption) - 3)
        x2.Caption = "{x2} = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [X2 SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option11.Value = True Then
    Answer.Caption = Round(CalcLPrime(TXTX2.Text, TXTV.Text), 3) & " lu"
    If assign.Value = 1 Then
        TXTX2PRIME.Text = Left(Answer.Caption, Len(Answer.Caption) - 3)
        x2prime.Caption = "{x2'} = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [X2' SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option12.Value = True Then
    Answer.Caption = Round(CalcDeltaT(TXTT2PRIME.Text, TXTV.Text), 3) & " tu"
    If assign.Value = 1 Then
        TXTT2.Text = Left(Answer.Caption, Len(Answer.Caption) - 3)
        t2.Caption = "{t2} = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [t2 SET TO " + Answer.Caption + "]", 0)
    End If
End If

If Option13.Value = True Then
    Answer.Caption = Round(CalcDeltaTPrime(TXTT2.Text, TXTV.Text), 3) & " tu"
    If assign.Value = 1 Then
        TXTT2PRIME.Text = Left(Answer.Caption, Len(Answer.Caption) - 3)
        t2prime.Caption = "{t2'} = " & Answer.Caption
        Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [T2' SET TO " + Answer.Caption + "]", 0)
    End If
End If
End Sub

Private Sub Command4_Click()
Form1.Show vbModal
End Sub

Private Sub Command7_Click()
Call uxprime_Click
End Sub

Private Sub Command8_Click()
Call x2prime_Click
End Sub

Private Sub Label23_Click()

End Sub

Private Sub Command9_Click()
Call x2_Click
End Sub

Private Sub Missing_DblClick()
Select Case Missing.List(Missing.ListIndex)
    Case "v"
        Call v_Click
    Case "ux"
        Call ux_Click
    Case "ux'"
        Call uxprime_Click
    Case "x2"
        Call x2_Click
    Case "x2'"
        Call x2prime_Click
    Case "t2"
        Call t2_Click
    Case "t2'"
        Call t2prime_Click
End Select
End Sub

Private Sub mnuFileAbout_Click()
MsgBox "Created by JJ Geewax on 10/12/03 for AP Physics."
End Sub

Private Sub mnuFileExit_Click()
If MsgBox("Are you sure you want to exit?", vbYesNo) = vbNo Then Exit Sub
End
End Sub

Private Sub mnuFileNew_Click()
lstSave.Clear
Call Command2_Click
Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [NEW FILE CREATED!]", 0)
End Sub

Private Sub mnuFileOpen_Click()
lstSave.Clear
CD.ShowOpen
If Not CD.FileName = "" Then
    Call Loadlist(CD.FileName, lstSave)
    TXTV.Text = lstSave.List(0)
    v.Caption = "v = " + TXTV.Text
    TXTUX.Text = lstSave.List(1)
    ux.Caption = "{ux} = " + TXTUX.Text
    TXTUXPRIME.Text = lstSave.List(2)
    uxprime.Caption = "{ux'} = " + TXTUXPRIME.Text
    TXTX2.Text = lstSave.List(3)
    x2.Caption = "{x2} = " + TXTX2.Text
    TXTX2PRIME.Text = lstSave.List(4)
    x2prime.Caption = "{x2'} = " + TXTX2PRIME.Text
    TXTT2.Text = lstSave.List(5)
    t2.Caption = "{t2} = " + TXTT2.Text
    TXTT2PRIME.Text = lstSave.List(6)
    t2prime.Caption = "{t2'} = " + TXTT2PRIME.Text
    Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [FILE LOADED!]", 0)
End If
End Sub

Private Sub mnuFileSave_Click()
On Error GoTo err:
lstSave.Clear
Call lstSave.AddItem(TXTV.Text, 0)
Call lstSave.AddItem(TXTUX.Text, 1)
Call lstSave.AddItem(TXTUXPRIME.Text, 2)
Call lstSave.AddItem(TXTX2.Text, 3)
Call lstSave.AddItem(TXTX2PRIME.Text, 4)
Call lstSave.AddItem(TXTT2.Text, 5)
Call lstSave.AddItem(TXTT2PRIME.Text, 6)
Call CD.ShowSave
If CD.FileName = "" Then
Exit Sub
Else
    Call Module1.SaveList(CD.FileName, lstSave)
    Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [FILE SAVED!]", 0)
    MsgBox CD.FileName
End If
Exit Sub
err:
Exit Sub
End Sub

Private Sub mnuHelpAbout_Click()
MsgBox "Created by JJ Geewax for AP Physics" + vbNewLine + "      Version 1.2      "
End Sub

Private Sub mnuHelpLenContr_Click()
Form4.Show vbModal
End Sub

Private Sub mnuHelpTimeDil_Click()
Form5.Show vbModal
End Sub

Private Sub Option1_Click()
Command3.Enabled = True
Uses.Clear
Missing.Clear

Uses.AddItem ("ux'")
Uses.AddItem ("ux")

If TXTUXPRIME.Text = "" Then Missing.AddItem ("ux'")
If TXTUX.Text = "" Then Missing.AddItem ("ux")

End Sub

Private Sub Option10_Click()
Command3.Enabled = True
Uses.Clear
Missing.Clear

Uses.AddItem ("x2'")
Uses.AddItem ("v")

If TXTX2PRIME.Text = "" Then Missing.AddItem ("x2'")
If TXTV.Text = "" Then Missing.AddItem ("v")

End Sub

Private Sub Option11_Click()
Command3.Enabled = True
Uses.Clear
Missing.Clear

Uses.AddItem ("x2")
Uses.AddItem ("v")

If TXTX2.Text = "" Then Missing.AddItem ("x2")
If TXTV.Text = "" Then Missing.AddItem ("v")

End Sub

Private Sub Option12_Click()
Command3.Enabled = True
Uses.Clear
Missing.Clear

Uses.AddItem ("t2'")
Uses.AddItem ("v")

If TXTT2PRIME.Text = "" Then Missing.AddItem ("t2'")
If TXTV.Text = "" Then Missing.AddItem ("v")

End Sub

Private Sub Option13_Click()
Command3.Enabled = True
Uses.Clear
Missing.Clear

Uses.AddItem ("t2")
Uses.AddItem ("v")

If TXTT2.Text = "" Then Missing.AddItem ("t2")
If TXTV.Text = "" Then Missing.AddItem ("v")

End Sub

Private Sub Option2_Click()
Command3.Enabled = True
Uses.Clear
Missing.Clear

Uses.AddItem ("x2")
Uses.AddItem ("x2'")

If TXTX2.Text = "" Then Missing.AddItem ("x2")
If TXTX2PRIME.Text = "" Then Missing.AddItem ("x2'")

End Sub

Private Sub Option3_Click()
Command3.Enabled = True
Uses.Clear
Missing.Clear

Uses.AddItem ("ux'")
Uses.AddItem ("v")

If TXTUXPRIME.Text = "" Then Missing.AddItem ("ux'")
If TXTV.Text = "" Then Missing.AddItem ("v")

End Sub

Private Sub Option4_Click()
Command3.Enabled = True
Missing.Clear
Uses.Clear

Uses.AddItem ("ux")
Uses.AddItem ("v")

If TXTUX.Text = "" Then Missing.AddItem ("ux")
If TXTV.Text = "" Then Missing.AddItem ("v")

End Sub

Private Sub Option5_Click()
Command3.Enabled = True
Missing.Clear
Uses.Clear

Uses.AddItem ("x2'")
Uses.AddItem ("v")
Uses.AddItem ("t2'")

If TXTX2PRIME.Text = "" Then Missing.AddItem ("x2'")
If TXTV.Text = "" Then Missing.AddItem ("v")
If TXTT2PRIME.Text = "" Then Missing.AddItem ("t2'")
End Sub

Private Sub Option6_Click()

Command3.Enabled = True
Missing.Clear
Uses.Clear

Uses.AddItem ("x2")
Uses.AddItem ("v")
Uses.AddItem ("t2")

If TXTX2.Text = "" Then Missing.AddItem ("x2")
If TXTV.Text = "" Then Missing.AddItem ("v")
If TXTT2.Text = "" Then Missing.AddItem ("t2")

End Sub

Private Sub Option7_Click()
Command3.Enabled = True
Missing.Clear
Uses.Clear

Uses.AddItem ("t2'")
Uses.AddItem ("v")
Uses.AddItem ("x2'")

If TXTT2PRIME.Text = "" Then Missing.AddItem ("t2'")
If TXTV.Text = "" Then Missing.AddItem ("v")
If TXTX2PRIME.Text = "" Then Missing.AddItem ("x2'")

End Sub

Private Sub Option8_Click()
Command3.Enabled = True
Missing.Clear
Uses.Clear

Uses.AddItem ("t2")
Uses.AddItem ("v")
Uses.AddItem ("x2")

If TXTT2.Text = "" Then Missing.AddItem ("t2")
If TXTV.Text = "" Then Missing.AddItem ("v")
If TXTX2.Text = "" Then Missing.AddItem ("x2")

End Sub

Private Sub Option9_Click()
Command3.Enabled = True

Uses.Clear
Missing.Clear

Call Uses.AddItem("t2")
Call Uses.AddItem("t2'")

If TXTT2.Text = "" Then Missing.AddItem ("t2")
If TXTT2PRIME.Text = "" Then Missing.AddItem ("t2'")

End Sub

Private Sub t2_Click()
Dim t21 As String
t21 = InputBox("What value do you want for T2 (in time-units)?", "Special Relativity")
If t21 = "" Then Exit Sub
t2.Caption = "{t2} = " + t21 + " tu"
TXTT2.Text = t21
Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [T2 SET TO " + t21 + " time-units]", 0)
End Sub

Private Sub t2prime_Click()
Dim t2prime1 As String
t2prime1 = InputBox("What value do you want for T2' (in time-units)?", "Special Relativity")
If t2prime1 = "" Then Exit Sub
t2prime.Caption = "{t2'} = " + t2prime1 + " tu"
TXTT2PRIME.Text = t2prime1
Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [T2' SET TO " + t2prime1 + " time-units]", 0)
End Sub

Private Sub TXTT2_Change()
Dim i As Integer
For i = 0 To Missing.ListCount - 1
    If Missing.List(i) = "t2" Then
        Missing.RemoveItem (i)
        Exit Sub
    End If
Next i
End Sub

Private Sub TXTT2_DblClick()
Call Command13_Click
End Sub

Private Sub TXTT2PRIME_Change()
Dim i As Integer
For i = 0 To Missing.ListCount - 1
    If Missing.List(i) = "t2'" Then
        Missing.RemoveItem (i)
        Exit Sub
    End If
Next i
End Sub

Private Sub TXTT2PRIME_DblClick()
Call Command12_Click
End Sub

Private Sub TXTUX_Change()
Dim i As Integer
For i = 0 To Missing.ListCount - 1
    If Missing.List(i) = "ux" Then
        Missing.RemoveItem (i)
        Exit Sub
    End If
Next i
End Sub

Private Sub TXTUX_DblClick()
Call Command10_Click
End Sub

Private Sub TXTUXPRIME_Change()
Dim i As Integer
For i = 0 To Missing.ListCount - 1
    If Missing.List(i) = "ux'" Then
        Missing.RemoveItem (i)
        Exit Sub
    End If
Next i
End Sub

Private Sub TXTUXPRIME_DblClick()
Call Command7_Click
End Sub

Private Sub TXTV_Change()
Dim i As Integer
For i = 0 To Missing.ListCount - 1
    If Missing.List(i) = "v" Then
        Missing.RemoveItem (i)
        Exit Sub
    End If
Next i
End Sub

Private Sub TXTV_DblClick()
Call Command11_Click
End Sub

Private Sub TXTX2_Change()
Dim i As Integer
For i = 0 To Missing.ListCount - 1
    If Missing.List(i) = "x2" Then
        Missing.RemoveItem (i)
        Exit Sub
    End If
Next i
End Sub

Private Sub TXTX2_DblClick()
Call Command9_Click
End Sub

Private Sub TXTX2PRIME_Change()
Dim i As Integer
For i = 0 To Missing.ListCount - 1
    If Missing.List(i) = "x2'" Then
        Missing.RemoveItem (i)
        Exit Sub
    End If
Next i
End Sub

Private Sub TXTX2PRIME_DblClick()
Call Command8_Click
End Sub

Private Sub ux_Click()
Dim ux1 As String
restart:
ux1 = InputBox("What multiple of c do you want to assign to ux? Leave blank to cancel.", "Special Relativity")
If ux1 = "" Then Exit Sub
If Val(ux1) >= 1 Then
    MsgBox "You cannot go faster or equal to light!"
    ux1 = ""
    GoTo restart:
End If
TXTUX.Text = ux1
ux.Caption = "{ux} =" + ux1 + "c"
Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [UX SET TO " + ux1 + "c]", 0)
End Sub

Private Sub uxprime_Click()
Dim uxprime1 As String
restart:
uxprime1 = InputBox("What multiple of c do you want to assign to ux'? Leave blank to cancel.", "Special Relativity")
If uxprime1 = "" Then Exit Sub
If Val(uxprime1) >= 1 Then
    MsgBox "You cannot go faster or equal to light!"
    uxprime1 = ""
    GoTo restart:
End If
uxprime.Caption = "{ux'} =" + uxprime1 + "c"
TXTUXPRIME.Text = uxprime1
Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [UX' SET TO " + uxprime1 + "c]", 0)
End Sub

Private Sub v_Click()
Dim v1 As String
restart:
v1 = InputBox("What multiple of c do you want to assign to v? Leave blank to cancel.", "Special Relativity")
If v1 = "" Then Exit Sub
If Val(v1) >= 1 Then
    MsgBox "You cannot go faster or equal to light!"
    v1 = ""
    GoTo restart:
End If
v.Caption = "v = " + v1 + "c"
TXTV.Text = v1
Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [V SET TO " + v1 + "c]", 0)
End Sub

Private Sub x2_Click()
Dim x21 As String
x21 = InputBox("What value do you want to assign to x2 in terms of light-units? Leave blank to cancel.", "Special Relativity")
If x21 = "" Then Exit Sub
x2.Caption = "{x2} = " + x21 + " lu"
TXTX2.Text = x21
Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [X2 SET TO " + x21 + " light-units]", 0)
End Sub

Private Sub x2prime_Click()
Dim x2prime1 As String
x2prime1 = InputBox("What value do you want to assign to x2' in terms of light-units? Leave blank to cancel.", "Special Relativity")
If x2prime1 = "" Then Exit Sub
x2prime.Caption = "{x2'} = " + x2prime1 + " lu"
TXTX2PRIME.Text = x2prime1
Call Status.AddItem(Format(DateTime.Now, "hh:mm:ss") + " - [X2' SET TO " + x2prime1 + " light-units]", 0)
End Sub
