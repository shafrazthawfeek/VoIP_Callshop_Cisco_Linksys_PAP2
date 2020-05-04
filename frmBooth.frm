VERSION 5.00
Begin VB.Form frmBooth 
   Caption         =   "Booth Management"
   ClientHeight    =   8610
   ClientLeft      =   6915
   ClientTop       =   2160
   ClientWidth     =   12795
   Icon            =   "frmBooth.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   12795
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbBoothType 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmBooth.frx":0442
      Left            =   1815
      List            =   "frmBooth.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   100
      Top             =   180
      Width           =   1290
   End
   Begin VB.PictureBox canvas 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   120
      ScaleHeight     =   7695
      ScaleWidth      =   12615
      TabIndex        =   1
      Top             =   720
      Width           =   12615
      Begin VB.PictureBox container 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   12000
         Left            =   50
         ScaleHeight     =   12000
         ScaleWidth      =   12255
         TabIndex        =   3
         Top             =   0
         Width           =   12255
         Begin VB.Frame frmBooth 
            Caption         =   "Booth15"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   14
            Left            =   8235
            TabIndex        =   94
            Top             =   7680
            Width           =   3885
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   14
               ItemData        =   "frmBooth.frx":0463
               Left            =   1455
               List            =   "frmBooth.frx":046D
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   1305
               Width           =   1080
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   14
               Left            =   165
               TabIndex        =   96
               Top             =   375
               Width           =   2445
            End
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   14
               Left            =   1305
               TabIndex        =   95
               Top             =   765
               Width           =   2055
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   14
               Left            =   150
               TabIndex        =   99
               Top             =   1350
               Width           =   1350
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   14
               Left            =   150
               TabIndex        =   98
               Top             =   855
               Width           =   1110
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth14"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   13
            Left            =   4125
            TabIndex        =   88
            Top             =   7680
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   13
               Left            =   1305
               TabIndex        =   91
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   13
               Left            =   165
               TabIndex        =   90
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   13
               ItemData        =   "frmBooth.frx":0477
               Left            =   1455
               List            =   "frmBooth.frx":0481
               Style           =   2  'Dropdown List
               TabIndex        =   89
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   13
               Left            =   150
               TabIndex        =   93
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   13
               Left            =   150
               TabIndex        =   92
               Top             =   1350
               Width           =   1350
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth13"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   12
            Left            =   0
            TabIndex        =   82
            Top             =   7680
            Width           =   3885
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   12
               ItemData        =   "frmBooth.frx":048B
               Left            =   1455
               List            =   "frmBooth.frx":0495
               Style           =   2  'Dropdown List
               TabIndex        =   85
               Top             =   1305
               Width           =   1080
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   12
               Left            =   165
               TabIndex        =   84
               Top             =   375
               Width           =   2445
            End
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   12
               Left            =   1305
               TabIndex        =   83
               Top             =   765
               Width           =   2055
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   12
               Left            =   150
               TabIndex        =   87
               Top             =   1350
               Width           =   1350
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   12
               Left            =   150
               TabIndex        =   86
               Top             =   855
               Width           =   1110
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth12"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   11
            Left            =   8235
            TabIndex        =   76
            Top             =   5760
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   11
               Left            =   1305
               TabIndex        =   79
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   11
               Left            =   165
               TabIndex        =   78
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   11
               ItemData        =   "frmBooth.frx":049F
               Left            =   1455
               List            =   "frmBooth.frx":04A9
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   11
               Left            =   150
               TabIndex        =   81
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   11
               Left            =   150
               TabIndex        =   80
               Top             =   1350
               Width           =   1350
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth11"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   10
            Left            =   4125
            TabIndex        =   70
            Top             =   5760
            Width           =   3885
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   10
               ItemData        =   "frmBooth.frx":04B3
               Left            =   1455
               List            =   "frmBooth.frx":04BD
               Style           =   2  'Dropdown List
               TabIndex        =   73
               Top             =   1305
               Width           =   1080
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   10
               Left            =   165
               TabIndex        =   72
               Top             =   375
               Width           =   2445
            End
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   10
               Left            =   1305
               TabIndex        =   71
               Top             =   765
               Width           =   2055
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   10
               Left            =   150
               TabIndex        =   75
               Top             =   1350
               Width           =   1350
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   10
               Left            =   150
               TabIndex        =   74
               Top             =   855
               Width           =   1110
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   2
            Left            =   8235
            TabIndex        =   64
            Top             =   0
            Width           =   3885
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   2
               ItemData        =   "frmBooth.frx":04C7
               Left            =   1455
               List            =   "frmBooth.frx":04D1
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   1305
               Width           =   1080
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   165
               TabIndex        =   66
               Top             =   375
               Width           =   2445
            End
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   2
               Left            =   1305
               TabIndex        =   65
               Top             =   765
               Width           =   2055
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   150
               TabIndex        =   69
               Top             =   1350
               Width           =   1350
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   150
               TabIndex        =   68
               Top             =   855
               Width           =   1110
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth10"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   9
            Left            =   0
            TabIndex        =   58
            Top             =   5760
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   9
               Left            =   1305
               TabIndex        =   61
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   9
               Left            =   165
               TabIndex        =   60
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   9
               ItemData        =   "frmBooth.frx":04DB
               Left            =   1455
               List            =   "frmBooth.frx":04E5
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   9
               Left            =   150
               TabIndex        =   63
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   9
               Left            =   150
               TabIndex        =   62
               Top             =   1350
               Width           =   1350
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth9"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   8
            Left            =   8235
            TabIndex        =   52
            Top             =   3840
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   8
               Left            =   1305
               TabIndex        =   55
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   8
               Left            =   165
               TabIndex        =   54
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   8
               ItemData        =   "frmBooth.frx":04EF
               Left            =   1455
               List            =   "frmBooth.frx":04F9
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   8
               Left            =   150
               TabIndex        =   57
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   8
               Left            =   150
               TabIndex        =   56
               Top             =   1350
               Width           =   1350
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth8"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   7
            Left            =   4125
            TabIndex        =   46
            Top             =   3840
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   7
               Left            =   1305
               TabIndex        =   49
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   7
               Left            =   165
               TabIndex        =   48
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   7
               ItemData        =   "frmBooth.frx":0503
               Left            =   1455
               List            =   "frmBooth.frx":050D
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   7
               Left            =   150
               TabIndex        =   51
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   7
               Left            =   150
               TabIndex        =   50
               Top             =   1350
               Width           =   1350
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth7"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   6
            Left            =   0
            TabIndex        =   40
            Top             =   3840
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   6
               Left            =   1305
               TabIndex        =   43
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   6
               Left            =   165
               TabIndex        =   42
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   6
               ItemData        =   "frmBooth.frx":0517
               Left            =   1455
               List            =   "frmBooth.frx":0521
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   6
               Left            =   150
               TabIndex        =   45
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   6
               Left            =   150
               TabIndex        =   44
               Top             =   1350
               Width           =   1350
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth6"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   5
            Left            =   8235
            TabIndex        =   34
            Top             =   1920
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   5
               Left            =   1305
               TabIndex        =   37
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   165
               TabIndex        =   36
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   5
               ItemData        =   "frmBooth.frx":052B
               Left            =   1455
               List            =   "frmBooth.frx":0535
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   150
               TabIndex        =   39
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   5
               Left            =   150
               TabIndex        =   38
               Top             =   1350
               Width           =   1350
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth5"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   4
            Left            =   4125
            TabIndex        =   28
            Top             =   1920
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   4
               Left            =   1305
               TabIndex        =   31
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   165
               TabIndex        =   30
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   4
               ItemData        =   "frmBooth.frx":053F
               Left            =   1455
               List            =   "frmBooth.frx":0549
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   4
               Left            =   150
               TabIndex        =   33
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   150
               TabIndex        =   32
               Top             =   1350
               Width           =   1350
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth4"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   3
            Left            =   0
            TabIndex        =   22
            Top             =   1920
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   3
               Left            =   1305
               TabIndex        =   25
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   165
               TabIndex        =   24
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   3
               ItemData        =   "frmBooth.frx":0553
               Left            =   1455
               List            =   "frmBooth.frx":055D
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   3
               Left            =   150
               TabIndex        =   27
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   150
               TabIndex        =   26
               Top             =   1350
               Width           =   1350
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   1
            Left            =   4125
            TabIndex        =   16
            Top             =   15
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   1
               Left            =   1305
               TabIndex        =   19
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   165
               TabIndex        =   18
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   1
               ItemData        =   "frmBooth.frx":0567
               Left            =   1455
               List            =   "frmBooth.frx":0571
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   150
               TabIndex        =   21
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   150
               TabIndex        =   20
               Top             =   1350
               Width           =   1350
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Top             =   30
            Width           =   3885
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   0
               ItemData        =   "frmBooth.frx":057B
               Left            =   1455
               List            =   "frmBooth.frx":0585
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   1305
               Width           =   1080
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   165
               TabIndex        =   12
               Top             =   375
               Width           =   2445
            End
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   0
               Left            =   1305
               TabIndex        =   11
               Top             =   765
               Width           =   2055
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   150
               TabIndex        =   15
               Top             =   1350
               Width           =   1350
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   150
               TabIndex        =   14
               Top             =   855
               Width           =   1110
            End
         End
         Begin VB.Frame frmBooth 
            Caption         =   "Booth16"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   15
            Left            =   0
            TabIndex        =   4
            Top             =   9600
            Width           =   3885
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   15
               Left            =   1305
               TabIndex        =   7
               Top             =   765
               Width           =   2055
            End
            Begin VB.CheckBox chkEnable 
               Caption         =   "Booth Enabled"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   15
               Left            =   165
               TabIndex        =   6
               Top             =   375
               Width           =   2445
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   15
               ItemData        =   "frmBooth.frx":058F
               Left            =   1455
               List            =   "frmBooth.frx":0599
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1305
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "IP Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   15
               Left            =   150
               TabIndex        =   9
               Top             =   855
               Width           =   1110
            End
            Begin VB.Label lblPort 
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   15
               Left            =   150
               TabIndex        =   8
               Top             =   1350
               Width           =   1350
            End
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1575
         LargeChange     =   2000
         Left            =   12300
         Max             =   6000
         SmallChange     =   1000
         TabIndex        =   2
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   1395
   End
End
Attribute VB_Name = "frmBooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub canvas_Resize()
    VScroll1.Left = canvas.Width - VScroll1.Width
    VScroll1.Height = canvas.Height - VScroll1.Top
    container.Width = canvas.Width - VScroll1.Width - container.Left
End Sub

Private Sub chkEnable_Click(Index As Integer)
Dim x As Integer
    
    'txtIP(Index).Enabled = Not txtIP(Index).Enabled
    'Combo1(Index).Enabled = Not Combo1(Index).Enabled
    
    If chkEnable(Index).Value = 1 Then
        If Index < 15 Then
            frmBooth(Index).Visible = True
            frmBooth(Index + 1).Visible = True
        End If
    
    Else
    
        For x = Index + 1 To 15
            frmBooth(x).Visible = False
            chkEnable(x).Value = 0
        Next
        
    End If
    
    
End Sub

Private Sub GetBoothConfig()
Dim bn As Integer

    Set rs = cn.Execute("select * from booths order by BoothNumber ASC;")
    If Not (rs.BOF) And Not (rs.EOF) Then
        rs.MoveFirst
    Else
        frmBooth(0).Visible = True
    End If
    
    While rs.EOF = False
        bn = rs!BoothNumber
        txtIP(bn - 1) = rs!IPAddress
        Combo1(bn - 1).ListIndex = rs!Port - 1
        chkEnable(bn - 1).Value = 1
        chkEnable_Click (bn - 1)
                
        rs.MoveNext
    Wend
    
    rs.Close
    
    Set rs = cn.Execute("select BoothIsSmall from Settings;")
    If Not (rs.BOF) And Not (rs.EOF) Then
        rs.MoveFirst
        If rs!BoothIsSmall = 0 Then
            cmbBoothType.ListIndex = 0
        Else
            cmbBoothType.ListIndex = 1
        End If
        
    Else
        cmbBoothType.ListIndex = 0
    End If
End Sub

Private Sub Command1_Click()
Dim a As Integer
Dim s As String
On Error Resume Next
        
    cn.Execute ("delete from booths;")
    
    For a = 0 To 15
        'If frmBooth(a).Visible = False Then Exit For
        If chkEnable(a).Value = 0 Then Exit For
        s = "Insert Into Booths Values(" & a + 1 & ",'" & txtIP(a) & "'," & Combo1(a).Text & ", True);"
        cn.Execute (s)
    Next
    
    If cmbBoothType.ListIndex = 0 Then
        cn.Execute ("Update Settings set BoothIsSmall=0")
    Else
        cn.Execute ("Update Settings set BoothIsSmall=1")
    End If
    
    If MsgBox("Booths Saved. Reload Booths now? All Calls currently in progress will be lost!" & vbCrLf & "You may decide to Reload later by restarting the Callshop application." & vbCrLf & "Press YES to reload now, NO to Reload Later", vbInformation + vbYesNo, "Saved") = vbYes Then
        
        Unload Form1
        Load Form1
        Form1.Show
    End If
    
    Unload Me
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Shift = 0 Then
    Select Case KeyCode
        Case 33, 34:
            If VScroll1.Visible Then VScroll1.SetFocus
    
    End Select
    End If
    
End Sub


Private Sub Form_Load()
Dim a As Integer

    OpenDB
   
    For a = 0 To 15
        frmBooth(a).Visible = False
    Next
    
    GetBoothConfig
      
End Sub

Private Sub OpenDB()
    Set cn = New ADODB.Connection
    cn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\callshop.cfg"
   
End Sub

Private Sub CloseDB()
    cn.Close
    
    Set cn = Nothing
    Set rs = Nothing
    
End Sub


Private Sub Form_Resize()
    canvas.Width = Me.ScaleWidth - canvas.Left
    canvas.Height = Me.ScaleHeight - canvas.Top
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseDB

End Sub

Private Sub VScroll1_Change()
    VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
    container.Top = -VScroll1.Value
End Sub
