VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Information Window"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Illustration"
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   4815
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ux'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   720
         Width           =   255
      End
      Begin VB.Line Line20 
         X1              =   2520
         X2              =   2655
         Y1              =   840
         Y2              =   960
      End
      Begin VB.Line Line19 
         X1              =   2520
         X2              =   2655
         Y1              =   1095
         Y2              =   960
      End
      Begin VB.Line Line18 
         X1              =   2160
         X2              =   2640
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "v"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         Top             =   960
         Width           =   255
      End
      Begin VB.Line Line17 
         X1              =   4200
         X2              =   4335
         Y1              =   1080
         Y2              =   1200
      End
      Begin VB.Line Line16 
         X1              =   4200
         X2              =   4335
         Y1              =   1335
         Y2              =   1200
      End
      Begin VB.Line Line15 
         X1              =   3840
         X2              =   4320
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Shape Shape6 
         Height          =   255
         Left            =   1920
         Shape           =   3  'Circle
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CS"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CS'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "y'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   135
      End
      Begin VB.Line Line14 
         X1              =   1080
         X2              =   3720
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line13 
         X1              =   1680
         X2              =   1680
         Y1              =   600
         Y2              =   1800
      End
      Begin VB.Shape Shape5 
         Height          =   1455
         Left            =   960
         Shape           =   2  'Oval
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "y"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   12
         Top             =   2280
         Width           =   135
      End
      Begin VB.Line Line12 
         X1              =   120
         X2              =   4680
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line11 
         X1              =   600
         X2              =   600
         Y1              =   240
         Y2              =   2760
      End
      Begin VB.Shape Shape4 
         Height          =   2535
         Left            =   120
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Transformations and Definitions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.Frame Frame11 
         Caption         =   "Velocity Trans"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   3960
         TabIndex        =   9
         Top             =   240
         Width           =   1575
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   120
            Picture         =   "Form1.frx":0000
            ScaleHeight     =   1905
            ScaleWidth      =   1305
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Definitions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1575
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   120
            Picture         =   "Form1.frx":8442
            ScaleHeight     =   1065
            ScaleWidth      =   1305
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Standard Conversions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   120
            Picture         =   "Form1.frx":C870
            ScaleHeight     =   1545
            ScaleWidth      =   1785
            TabIndex        =   6
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Time Trans."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   5640
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   120
            Picture         =   "Form1.frx":14AA2
            ScaleHeight     =   1305
            ScaleWidth      =   1305
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Position Trans."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2280
         TabIndex        =   1
         Top             =   1800
         Width           =   1575
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            Picture         =   "Form1.frx":19FE4
            ScaleHeight     =   705
            ScaleWidth      =   1305
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Label Label7 
      Caption         =   "{x'} = x-position in CS' (moving bubble)."
      Height          =   255
      Left            =   5040
      TabIndex        =   26
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "{x} = x-position in CS (ground)."
      Height          =   255
      Left            =   5040
      TabIndex        =   25
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "{t} = time in CS (ground)."
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "{t'} = time in CS' (moving bubble)."
      Height          =   255
      Left            =   5040
      TabIndex        =   23
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "{v} = velocity of CS' (moving bubble) relative to CS (ground)."
      Height          =   495
      Left            =   5040
      TabIndex        =   22
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "{ux'} = velocity of moving object in CS' relative to CS' (moving bubble)."
      Height          =   495
      Left            =   5040
      TabIndex        =   21
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "{ux} = velocity of moving object in CS' relative to CS (ground)."
      Height          =   495
      Left            =   5040
      TabIndex        =   20
      Top             =   3360
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
