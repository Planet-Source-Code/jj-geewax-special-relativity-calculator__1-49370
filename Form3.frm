VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Other Relativity Problems"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form3"
   ScaleHeight     =   6510
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame9 
      Caption         =   "Standard Conversions"
      Height          =   1935
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   2055
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         Picture         =   "Form3.frx":0000
         ScaleHeight     =   1545
         ScaleWidth      =   1785
         TabIndex        =   35
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox TXTX1PRIME 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TXTX1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set"
      Height          =   255
      Left            =   14760
      TabIndex        =   25
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   255
      Left            =   14760
      TabIndex        =   24
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TXTV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TXTX2PRIME 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TXTX2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   7920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Set"
      Height          =   255
      Left            =   14760
      TabIndex        =   14
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Set"
      Height          =   255
      Left            =   14760
      TabIndex        =   13
      Top             =   7920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Set"
      Height          =   255
      Left            =   14760
      TabIndex        =   12
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Illustration for Length Contraction"
      Height          =   2895
      Left            =   7920
      TabIndex        =   0
      Top             =   6720
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Shape Shape2 
         BorderStyle     =   3  'Dot
         Height          =   135
         Left            =   1200
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "{L}"
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "{L'}"
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label x2 
         BackStyle       =   0  'Transparent
         Caption         =   "{x2}"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label x1 
         BackStyle       =   0  'Transparent
         Caption         =   "{x1}"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label x1prime 
         BackStyle       =   0  'Transparent
         Caption         =   "{x1'}"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label x2prime 
         BackStyle       =   0  'Transparent
         Caption         =   "{x2'}"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   1200
         X2              =   1200
         Y1              =   240
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   2520
         X2              =   2520
         Y1              =   240
         Y2              =   2280
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         Height          =   135
         Left            =   1200
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Shape Shape4 
         Height          =   2535
         Left            =   120
         Top             =   240
         Width           =   4575
      End
      Begin VB.Line Line11 
         X1              =   480
         X2              =   480
         Y1              =   240
         Y2              =   2760
      End
      Begin VB.Line Line12 
         X1              =   120
         X2              =   4680
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "y"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   135
      End
      Begin VB.Shape Shape5 
         Height          =   1455
         Left            =   480
         Shape           =   2  'Oval
         Top             =   480
         Width           =   3255
      End
      Begin VB.Line Line13 
         X1              =   1200
         X2              =   1200
         Y1              =   600
         Y2              =   1800
      End
      Begin VB.Line Line14 
         X1              =   600
         X2              =   3600
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "y'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CS'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CS"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2280
         Width           =   255
      End
      Begin VB.Line Line15 
         X1              =   3720
         X2              =   4560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line16 
         X1              =   4440
         X2              =   4575
         Y1              =   1335
         Y2              =   1200
      End
      Begin VB.Line Line17 
         X1              =   4440
         X2              =   4575
         Y1              =   1080
         Y2              =   1200
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "{v}"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Label Label21 
      Caption         =   "x1' = "
      Height          =   255
      Left            =   12840
      TabIndex        =   31
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label20 
      Caption         =   "x1 = "
      Height          =   255
      Left            =   12840
      TabIndex        =   30
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "light - units"
      Height          =   255
      Left            =   13680
      TabIndex        =   29
      Top             =   7560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "light - units"
      Height          =   255
      Left            =   13680
      TabIndex        =   28
      Top             =   7200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "v = "
      Height          =   255
      Left            =   12840
      TabIndex        =   23
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "x2' = "
      Height          =   255
      Left            =   12840
      TabIndex        =   22
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "x2 = "
      Height          =   255
      Left            =   12840
      TabIndex        =   21
      Top             =   7920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "light - units"
      Height          =   255
      Left            =   13680
      TabIndex        =   20
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "light - units"
      Height          =   255
      Left            =   13680
      TabIndex        =   19
      Top             =   7920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "c"
      Height          =   255
      Left            =   13680
      TabIndex        =   18
      Top             =   6840
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label4_Click()

End Sub
