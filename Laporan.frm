VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Laporan 
   Caption         =   "Laporan"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13995
   BeginProperty Font 
      Name            =   "Century"
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
   ScaleHeight     =   5055
   ScaleWidth      =   13995
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   12120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   11
      TabsPerRow      =   11
      TabHeight       =   1058
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Master"
      TabPicture(0)   =   "Laporan.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "List1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Kas"
      TabPicture(1)   =   "Laporan.frx":08DA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Laporan.frx":11B4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Penjualan"
      TabPicture(3)   =   "Laporan.frx":11D0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame6"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Retur Penjualan"
      TabPicture(4)   =   "Laporan.frx":1AAA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame8"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame9"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "Laporan.frx":2384
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Pembelian"
      TabPicture(6)   =   "Laporan.frx":23A0
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame10"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame11"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Frame12"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Retur Pembelian"
      TabPicture(7)   =   "Laporan.frx":2C7A
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame15"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Frame14"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Frame13"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "Tab 8"
      TabPicture(8)   =   "Laporan.frx":3554
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Hutang"
      TabPicture(9)   =   "Laporan.frx":3570
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame16"
      Tab(9).Control(1)=   "Frame17"
      Tab(9).Control(2)=   "Frame18"
      Tab(9).ControlCount=   3
      TabCaption(10)  =   "Piutang"
      TabPicture(10)  =   "Laporan.frx":3E4A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame21"
      Tab(10).Control(1)=   "Frame20"
      Tab(10).Control(2)=   "Frame19"
      Tab(10).ControlCount=   3
      Begin VB.Frame Frame21 
         Caption         =   "Bulanan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -65880
         TabIndex        =   109
         Top             =   3360
         Width           =   4500
         Begin VB.ComboBox Combo35 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   112
            Top             =   360
            Width           =   1500
         End
         Begin VB.ComboBox Combo34 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   111
            Top             =   720
            Width           =   1500
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   110
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label36 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tahun"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   114
            Top             =   720
            Width           =   1250
         End
         Begin VB.Label Label35 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bulan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   113
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Harian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -65880
         TabIndex        =   105
         Top             =   840
         Width           =   4500
         Begin VB.ComboBox Combo33 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   107
            Top             =   360
            Width           =   1500
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   106
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label34 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Mingguan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -65880
         TabIndex        =   99
         Top             =   1920
         Width           =   4500
         Begin VB.ComboBox Combo32 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   102
            Top             =   360
            Width           =   1500
         End
         Begin VB.ComboBox Combo31 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   101
            Top             =   720
            Width           =   1500
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   100
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label33 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Awal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   104
            Top             =   360
            Width           =   1250
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Akhir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   103
            Top             =   720
            Width           =   1250
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Mingguan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -66960
         TabIndex        =   93
         Top             =   1920
         Width           =   4500
         Begin VB.CommandButton Command18 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   96
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox Combo30 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   95
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox Combo29 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   94
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label31 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Akhir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   98
            Top             =   720
            Width           =   1250
         End
         Begin VB.Label Label30 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Awal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   97
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Harian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -66960
         TabIndex        =   89
         Top             =   840
         Width           =   4500
         Begin VB.CommandButton Command17 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   91
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox Combo28 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   90
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label29 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Bulanan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -66960
         TabIndex        =   83
         Top             =   3360
         Width           =   4500
         Begin VB.CommandButton Command16 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   86
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox Combo27 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   85
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox Combo26 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   84
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label28 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bulan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   88
            Top             =   360
            Width           =   1250
         End
         Begin VB.Label Label27 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tahun"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   87
            Top             =   720
            Width           =   1250
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Bulanan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -69000
         TabIndex        =   77
         Top             =   3360
         Width           =   4500
         Begin VB.CommandButton Command15 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   80
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox Combo25 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   79
            Top             =   360
            Width           =   1500
         End
         Begin VB.ComboBox Combo24 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   78
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label26 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tahun"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   82
            Top             =   720
            Width           =   1250
         End
         Begin VB.Label Label25 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bulan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Harian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -69000
         TabIndex        =   73
         Top             =   840
         Width           =   4500
         Begin VB.CommandButton Command14 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   75
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox Combo23 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   74
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label24 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   76
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Mingguan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -69000
         TabIndex        =   67
         Top             =   1920
         Width           =   4500
         Begin VB.CommandButton Command13 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   70
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox Combo22 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   69
            Top             =   360
            Width           =   1500
         End
         Begin VB.ComboBox Combo21 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   68
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label23 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Awal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   1250
         End
         Begin VB.Label Label22 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Akhir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   71
            Top             =   720
            Width           =   1250
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Bulanan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -70680
         TabIndex        =   61
         Top             =   3360
         Width           =   4500
         Begin VB.ComboBox Combo20 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   64
            Top             =   360
            Width           =   1500
         End
         Begin VB.ComboBox Combo19 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   63
            Top             =   720
            Width           =   1500
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   62
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label21 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tahun"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   66
            Top             =   720
            Width           =   1250
         End
         Begin VB.Label Label20 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bulan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Harian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -70680
         TabIndex        =   57
         Top             =   840
         Width           =   4500
         Begin VB.ComboBox Combo18 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   59
            Top             =   360
            Width           =   1500
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   58
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Mingguan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -70680
         TabIndex        =   51
         Top             =   1920
         Width           =   4500
         Begin VB.ComboBox Combo17 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   54
            Top             =   360
            Width           =   1500
         End
         Begin VB.ComboBox Combo16 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   53
            Top             =   720
            Width           =   1500
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   52
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Awal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1250
         End
         Begin VB.Label Label17 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Akhir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   55
            Top             =   720
            Width           =   1250
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Mingguan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -72480
         TabIndex        =   45
         Top             =   1920
         Width           =   4500
         Begin VB.ComboBox Combo15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   48
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox Combo14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   47
            Top             =   360
            Width           =   1500
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   46
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label16 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Akhir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   50
            Top             =   720
            Width           =   1250
         End
         Begin VB.Label Label15 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Awal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Harian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -72480
         TabIndex        =   41
         Top             =   840
         Width           =   4500
         Begin VB.ComboBox Combo13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   43
            Top             =   360
            Width           =   1500
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   42
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Bulanan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -72480
         TabIndex        =   35
         Top             =   3360
         Width           =   4500
         Begin VB.ComboBox Combo12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   38
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox Combo11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   37
            Top             =   360
            Width           =   1500
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   36
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label13 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bulan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1250
         End
         Begin VB.Label Label12 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tahun"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1250
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Mingguan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74160
         TabIndex        =   29
         Top             =   1920
         Width           =   4500
         Begin VB.CommandButton Command6 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   32
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox Combo10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   31
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox Combo9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   30
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label11 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Akhir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   1250
         End
         Begin VB.Label Label10 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Awal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Harian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74160
         TabIndex        =   25
         Top             =   840
         Width           =   4500
         Begin VB.CommandButton Command5 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox Combo8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   26
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Bulanan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74160
         TabIndex        =   19
         Top             =   3360
         Width           =   4500
         Begin VB.CommandButton Command4 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   22
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox Combo7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox Combo6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   20
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bulan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   1250
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tahun"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1250
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mingguan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74760
         TabIndex        =   13
         Top             =   1920
         Width           =   4500
         Begin VB.CommandButton Command2 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   16
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   15
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   14
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label6 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Akhir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1250
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal Awal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Harian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   9
         Top             =   840
         Width           =   4500
         Begin VB.CommandButton Command1 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   10
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Bulanan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74760
         TabIndex        =   3
         Top             =   3360
         Width           =   4500
         Begin VB.CommandButton Command3 
            Caption         =   "Cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   6
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox Combo5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   5
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox Combo4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bulan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1250
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tahun"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   1250
         End
      End
      Begin VB.ListBox List1 
         Height          =   1635
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Pilih Tabel :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
List1.AddItem "Barang"
List1.AddItem "Kasir"
List1.AddItem "Pelanggan"
List1.AddItem "Pemasok"
List1.AddItem "Jasa"
List1.AddItem "Mekanik"

'================== arus kas


'On Error Resume Next
Call BukaDB
RSKas.Open "Select Distinct tanggal From Kas order By 1", Conn
RSKas.Requery
Do Until RSKas.EOF
    Combo1.AddItem Format(RSKas!tanggal, "DD-MMM-YYYY")
    Combo2.AddItem Format(RSKas!tanggal, "YYYY ,MM, DD")
    Combo3.AddItem Format(RSKas!tanggal, "YYYY ,MM, DD")
    RSKas.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGL As New ADODB.Recordset
RSTGL.Open "select distinct month(Tanggal) as Bulan from Kas", Conn
Do While Not RSTGL.EOF
    Combo4.AddItem RSTGL!Bulan & Space(5) & MonthName(RSTGL!Bulan)
    RSTGL.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(Tanggal)  as Tahun from Kas", Conn
Do While Not RSTHN.EOF
    Combo5.AddItem RSTHN!Tahun
    RSTHN.MoveNext
Loop
Conn.Close

'====================laporan penjualan
'On Error Resume Next
Call BukaDB
RSPenjualan.Open "Select Distinct Tanggal From Penjualan order By 1", Conn
RSPenjualan.Requery
Do Until RSPenjualan.EOF
    Combo8.AddItem Format(RSPenjualan!tanggal, "DD-MMM-YYYY")
    Combo9.AddItem Format(RSPenjualan!tanggal, "YYYY ,MM, DD")
    Combo10.AddItem Format(RSPenjualan!tanggal, "YYYY ,MM, DD")
    RSPenjualan.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGLJual As New ADODB.Recordset
RSTGLJual.Open "select distinct month(Tanggal) as Bulan from Penjualan", Conn
Do While Not RSTGLJual.EOF
    Combo6.AddItem RSTGLJual!Bulan & Space(5) & MonthName(RSTGLJual!Bulan)
    RSTGLJual.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHNJual As New ADODB.Recordset
RSTHNJual.Open "select distinct year(Tanggal)  as Tahun from Penjualan", Conn
Do While Not RSTHNJual.EOF
    Combo7.AddItem RSTHNJual!Tahun
    RSTHNJual.MoveNext
Loop
Conn.Close
'=================== retur penjualan

Call BukaDB
RSReturJual.Open "Select Distinct TanggalRetur From ReturJual order By 1", Conn
RSReturJual.Requery
Do Until RSReturJual.EOF
    Combo13.AddItem Format(RSReturJual!TanggalRetur, "DD-MMM-YYYY")
    Combo14.AddItem Format(RSReturJual!TanggalRetur, "YYYY ,MM, DD")
    Combo15.AddItem Format(RSReturJual!TanggalRetur, "YYYY ,MM, DD")
    RSReturJual.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGLRetur As New ADODB.Recordset
RSTGLRetur.Open "select distinct month(TanggalRetur) as Bulan from ReturJual", Conn
Do While Not RSTGLRetur.EOF
    Combo11.AddItem RSTGLRetur!Bulan & Space(5) & MonthName(RSTGLRetur!Bulan)
    RSTGLRetur.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHNRetur As New ADODB.Recordset
RSTHNRetur.Open "select distinct year(TanggalRetur)  as Tahun from ReturJual", Conn
Do While Not RSTHNRetur.EOF
    Combo12.AddItem RSTHNRetur!Tahun
    RSTHNRetur.MoveNext
Loop
Conn.Close

'================pembelian
'On Error Resume Next
Call BukaDB
RSPembelian.Open "Select Distinct Tanggal From Pembelian order By 1", Conn
RSPembelian.Requery
Do Until RSPembelian.EOF
    Combo18.AddItem Format(RSPembelian!tanggal, "DD-MMM-YYYY")
    Combo17.AddItem Format(RSPembelian!tanggal, "YYYY ,MM, DD")
    Combo16.AddItem Format(RSPembelian!tanggal, "YYYY ,MM, DD")
    RSPembelian.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGLBeli As New ADODB.Recordset
RSTGLBeli.Open "select distinct month(Tanggal) as Bulan from Pembelian", Conn
Do While Not RSTGLBeli.EOF
    Combo20.AddItem RSTGLBeli!Bulan & Space(5) & MonthName(RSTGLBeli!Bulan)
    RSTGLBeli.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHNBeli As New ADODB.Recordset
RSTHNBeli.Open "select distinct year(Tanggal)  as Tahun from Pembelian", Conn
Do While Not RSTHNBeli.EOF
    Combo19.AddItem RSTHNBeli!Tahun
    RSTHNBeli.MoveNext
Loop
Conn.Close

'============================ retur pembelian

Call BukaDB
RSReturBeli.Open "Select Distinct TanggalRetur From ReturBeli order By 1", Conn
RSReturBeli.Requery
Do Until RSReturBeli.EOF
    Combo23.AddItem Format(RSReturBeli!TanggalRetur, "DD-MMM-YYYY")
    Combo22.AddItem Format(RSReturBeli!TanggalRetur, "YYYY ,MM, DD")
    Combo21.AddItem Format(RSReturBeli!TanggalRetur, "YYYY ,MM, DD")
    RSReturBeli.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGLReturBeli As New ADODB.Recordset
RSTGLReturBeli.Open "select distinct month(TanggalRetur) as Bulan from ReturBeli", Conn
Do While Not RSTGLReturBeli.EOF
    Combo25.AddItem RSTGLReturBeli!Bulan & Space(5) & MonthName(RSTGLReturBeli!Bulan)
    RSTGLReturBeli.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHNReturBeli As New ADODB.Recordset
RSTHNReturBeli.Open "select distinct year(TanggalRetur)  as Tahun from ReturBeli", Conn
Do While Not RSTHNReturBeli.EOF
    Combo24.AddItem RSTHNReturBeli!Tahun
    RSTHNReturBeli.MoveNext
Loop
Conn.Close

'==================hutang
Call BukaDB
RSPembelian.Open "Select Distinct JatuhTempo From Pembelian WHERE SISA<>0 order By 1", Conn
RSPembelian.Requery
Do Until RSPembelian.EOF
    Combo28.AddItem Format(RSPembelian!jatuhtempo, "DD-MMM-YYYY")
    Combo29.AddItem Format(RSPembelian!jatuhtempo, "YYYY ,MM, DD")
    Combo30.AddItem Format(RSPembelian!jatuhtempo, "YYYY ,MM, DD")
    RSPembelian.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGLHutang As New ADODB.Recordset
RSTGLHutang.Open "select distinct month(JatuhTempo) as Bulan from Pembelian where sisa<>0", Conn
Do While Not RSTGLHutang.EOF
    Combo26.AddItem RSTGLHutang!Bulan & Space(5) & MonthName(RSTGLHutang!Bulan)
    RSTGLHutang.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHNHutang As New ADODB.Recordset
RSTHNHutang.Open "select distinct year(JatuhTempo)  as Tahun from Pembelian where sisa<>0", Conn
Do While Not RSTHNHutang.EOF
    Combo27.AddItem RSTHNHutang!Tahun
    RSTHNHutang.MoveNext
Loop
Conn.Close

'==================piutang
Call BukaDB
RSPenjualan.Open "Select Distinct JatuhTempo From Penjualan WHERE SISA<>0 order By 1", Conn
RSPenjualan.Requery
Do Until RSPenjualan.EOF
    Combo33.AddItem Format(RSPenjualan!jatuhtempo, "DD-MMM-YYYY")
    Combo32.AddItem Format(RSPenjualan!jatuhtempo, "YYYY ,MM, DD")
    Combo31.AddItem Format(RSPenjualan!jatuhtempo, "YYYY ,MM, DD")
    RSPenjualan.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGLPiutang As New ADODB.Recordset
RSTGLPiutang.Open "select distinct month(JatuhTempo) as Bulan from Penjualan where sisa<>0", Conn
Do While Not RSTGLPiutang.EOF
    Combo35.AddItem RSTGLPiutang!Bulan & Space(5) & MonthName(RSTGLPiutang!Bulan)
    RSTGLPiutang.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHNPiutang As New ADODB.Recordset
RSTHNPiutang.Open "select distinct year(JatuhTempo)  as Tahun from Penjualan where sisa<>0", Conn
Do While Not RSTHNPiutang.EOF
    Combo34.AddItem RSTHNPiutang!Tahun
    RSTHNPiutang.MoveNext
Loop
Conn.Close

End Sub

Private Sub List1_Click()
Select Case List1.ListIndex
    Case 0
        CR.DataFiles(0) = App.Path & "\dbretail.mdb"
        CR.ReportFileName = App.Path & "\Lap Barang.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case 1
        CR.DataFiles(0) = App.Path & "\dbretail.mdb"
        CR.ReportFileName = App.Path & "\Lap Pemakai.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case 2
        CR.DataFiles(0) = App.Path & "\dbretail.mdb"
        CR.ReportFileName = App.Path & "\Lap Pelanggan.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case 3
        CR.DataFiles(0) = App.Path & "\dbretail.mdb"
        CR.ReportFileName = App.Path & "\Lap Pemasok.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case 4
        CR.DataFiles(0) = App.Path & "\dbretail.mdb"
        CR.ReportFileName = App.Path & "\Lap Jasa.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case 5
        CR.DataFiles(0) = App.Path & "\dbretail.mdb"
        CR.ReportFileName = App.Path & "\Lap Mekanik.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1

End Select

End Sub

'=================arus kas

Private Sub Command1_Click()
If Combo1 = "" Then
    MsgBox "Pilih tanggal.."
    Combo1.SetFocus
    Exit Sub
End If
    CR.SelectionFormula = "Totext({Kas.Tanggal})='" & CDate(Combo1) & "'"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo1, "dd-mmm-yyyy") & "'"
    'CR.Formulas(0) = "TGLAWAL='" & Combo1 & "'"
    CR.ReportFileName = App.Path & "\Lap arus kas.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command2_Click()
    If Combo2 = "" Then
        MsgBox "Tanggal awal kosong", , "Informasi"
        Combo2.SetFocus
        Exit Sub
    Else
        If Combo3 < Combo2 Or Combo2 > Combo3 Then
            MsgBox "Tanggal terbalik"
            Combo3.SetFocus
            Exit Sub
        ElseIf Combo3 = Combo2 Then
            MsgBox "pilih Tanggal yang berbeda"
            Combo3.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{Kas.Tanggal} in date (" & Combo2 & ") to date (" & Combo3 & ")"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo2, "dd-mmm-yyyy") & "'"
    CR.Formulas(1) = "TGLAKHIR='" & Format(Combo3, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap arus kas.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command3_Click()
    Call BukaDB
    RSKas.Open "select * from Kas where month(Tanggal)='" & Val(Left(Combo4, 2)) & "' and year(Tanggal)='" & (Combo5) & "'", Conn
    If RSKas.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    CR.SelectionFormula = "Month({Kas.Tanggal})=" & Val(Left(Combo4, 2)) & " and Year({Kas.Tanggal})=" & Val(Combo5.Text)
    CR.ReportFileName = App.Path & "\lap arus kas.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub


'=================== penjualan
Private Sub Command5_Click()
If Combo8 = "" Then
    MsgBox "Pilih tangal.."
    Combo8.SetFocus
    Exit Sub
End If
    
    CR.SelectionFormula = "Totext({Penjualan.Tanggal})='" & CDate(Combo8) & "'"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo8, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap Penjualan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command6_Click()
    If Combo9 = "" Then
        MsgBox "Tanggal awal kosong", , "Informasi"
        Combo9.SetFocus
        Exit Sub
    Else
        If Combo10 < Combo9 Or Combo9 > Combo10 Then
            MsgBox "Tanggal terbalik"
            Combo10.SetFocus
            Exit Sub
        ElseIf Combo10 = Combo9 Then
            MsgBox "pilih tanggal yang berbeda"
            Combo10.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{Penjualan.Tanggal} in date (" & Combo9 & ") to date (" & Combo10 & ")"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo9, "dd-mmm-yyyy") & "'"
    CR.Formulas(1) = "TGLAKHIR='" & Format(Combo10, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap Penjualan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub command4_Click()
    Call BukaDB
    RSPenjualan.Open "select * from Penjualan where month(tanggal)='" & Val(Left(Combo6, 2)) & "' and year(tanggal)='" & (Combo7) & "'", Conn
    If RSPenjualan.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo6.SetFocus
    End If
    CR.SelectionFormula = "Month({Penjualan.Tanggal})=" & Val(Left(Combo6, 2)) & " and Year({Penjualan.Tanggal})=" & Val(Combo7.Text)
    CR.ReportFileName = App.Path & "\Lap Penjualan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

'================ retur penjualan
Private Sub Command8_Click()
If Combo13 = "" Then
    MsgBox "Pilih tanggal.."
    Combo13.SetFocus
    Exit Sub
End If
    
    CR.SelectionFormula = "Totext({ReturJual.TanggalRetur})='" & CDate(Combo13) & "'"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo13, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap retur Penjualan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command9_Click()
    If Combo14 = "" Then
        MsgBox "TanggalRetur awal kosong", , "Informasi"
        Combo14.SetFocus
        Exit Sub
    Else
        If Combo15 < Combo14 Or Combo14 > Combo15 Then
            MsgBox "TanggalRetur terbalik"
            Combo15.SetFocus
            Exit Sub
        ElseIf Combo15 = Combo14 Then
            MsgBox "pilih TanggalRetur yang berbeda"
            Combo15.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{ReturJual.TanggalRetur} in date (" & Combo14 & ") to date (" & Combo15 & ")"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo14, "dd-mmm-yyyy") & "'"
    CR.Formulas(1) = "TGLAKHIR='" & Format(Combo15, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap retur Penjualan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command7_Click()
    Call BukaDB
    RSReturJual.Open "select * from ReturJual where month(TanggalRetur)='" & Val(Left(Combo11, 2)) & "' and year(TanggalRetur)='" & (Combo12) & "'", Conn
    If RSReturJual.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo11.SetFocus
    End If
    CR.SelectionFormula = "Month({ReturJual.TanggalRetur})=" & Val(Left(Combo11, 2)) & " and Year({ReturJual.TanggalRetur})=" & Val(Combo12.Text)
    CR.ReportFileName = App.Path & "\Lap retur Penjualan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

'=================pembelian

Private Sub Command11_Click()
If Combo18 = "" Then
    MsgBox "Pilih tangal.."
    Combo18.SetFocus
    Exit Sub
End If
    
    CR.SelectionFormula = "Totext({Pembelian.Tanggal})='" & CDate(Combo18) & "'"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo18, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap pembelian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command10_Click()
    If Combo17 = "" Then
        MsgBox "Tanggal awal kosong", , "Informasi"
        Combo17.SetFocus
        Exit Sub
    Else
        If Combo16 < Combo17 Or Combo17 > Combo16 Then
            MsgBox "Tanggal terbalik"
            Combo16.SetFocus
            Exit Sub
        ElseIf Combo16 = Combo17 Then
            MsgBox "pilih tanggal yang berbeda"
            Combo16.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{Pembelian.Tanggal} in date (" & Combo17 & ") to date (" & Combo16 & ")"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo17, "dd-mmm-yyyy") & "'"
    CR.Formulas(1) = "TGLAKHIR='" & Format(Combo16, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap pembelian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command12_Click()
    Call BukaDB
    RSPembelian.Open "select * from Pembelian where month(tanggal)='" & Val(Left(Combo20, 2)) & "' and year(tanggal)='" & (Combo19) & "'", Conn
    If RSPembelian.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo20.SetFocus
    End If
    CR.SelectionFormula = "Month({Pembelian.Tanggal})=" & Val(Left(Combo20, 2)) & " and Year({Pembelian.Tanggal})=" & Val(Combo19.Text)
    CR.ReportFileName = App.Path & "\Lap pembelian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

'======================================== retur pembelian

Private Sub Command14_Click()
If Combo23 = "" Then
    MsgBox "Pilih tanggal.."
    Combo23.SetFocus
    Exit Sub
End If
    
    CR.SelectionFormula = "Totext({ReturBeli.TanggalRetur})='" & CDate(Combo23) & "'"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo23, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap retur pembelian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command13_Click()
    If Combo22 = "" Then
        MsgBox "TanggalRetur awal kosong", , "Informasi"
        Combo22.SetFocus
        Exit Sub
    Else
        If Combo21 < Combo22 Or Combo22 > Combo21 Then
            MsgBox "TanggalRetur terbalik"
            Combo21.SetFocus
            Exit Sub
        ElseIf Combo21 = Combo22 Then
            MsgBox "pilih TanggalRetur yang berbeda"
            Combo21.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{ReturBeli.TanggalRetur} in date (" & Combo22 & ") to date (" & Combo21 & ")"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo22, "dd-mmm-yyyy") & "'"
    CR.Formulas(1) = "TGLAKHIR='" & Format(Combo21, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap retur pembelian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command15_Click()
    Call BukaDB
    RSReturBeli.Open "select * from ReturBeli where month(TanggalRetur)='" & Val(Left(Combo25, 2)) & "' and year(TanggalRetur)='" & (Combo24) & "'", Conn
    If RSReturBeli.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo25.SetFocus
    End If
    CR.SelectionFormula = "Month({ReturBeli.TanggalRetur})=" & Val(Left(Combo25, 2)) & " and Year({ReturBeli.TanggalRetur})=" & Val(Combo24.Text)
    CR.ReportFileName = App.Path & "\Lap retur pembelian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub


'================hutang
Private Sub Command17_Click()
If Combo28 = "" Then
    MsgBox "Pilih JatuhTempo.."
    Combo28.SetFocus
    Exit Sub
End If

    CR.SelectionFormula = "Totext({Pembelian.JatuhTempo})='" & CDate(Combo28) & "' AND {PEMBELIAN.SISA}<>0"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo28, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap hutang.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command18_Click()
    If Combo29 = "" Then
        MsgBox "JatuhTempo awal kosong", , "Informasi"
        Combo29.SetFocus
        Exit Sub
    Else
        If Combo30 < Combo29 Or Combo29 > Combo30 Then
            MsgBox "JatuhTempo terbalik"
            Combo30.SetFocus
            Exit Sub
        ElseIf Combo30 = Combo29 Then
            MsgBox "pilih JatuhTempo yang berbeda"
            Combo30.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{Pembelian.JatuhTempo} in date (" & Combo29 & ") to date (" & Combo30 & ") AND {PEMBELIAN.SISA}<>0"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo29, "dd-mmm-yyyy") & "'"
    CR.Formulas(1) = "TGLAKHIR='" & Format(Combo30, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap HUTANG.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command16_Click()
    Call BukaDB
    RSPembelian.Open "select * from Pembelian where month(JatuhTempo)='" & Val(Left(Combo26, 2)) & "' and year(JatuhTempo)='" & (Combo27) & "' and sisa<>0", Conn
    If RSPembelian.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo26.SetFocus
    End If
    CR.SelectionFormula = "Month({Pembelian.JatuhTempo})=" & Val(Left(Combo26, 2)) & " and Year({pembelian.JatuhTempo})=" & Val(Combo27.Text) & " and {pembelian.sisa}<>0"
    CR.ReportFileName = App.Path & "\Lap hutang.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub


'================piutang
Private Sub Command20_Click()
If Combo33 = "" Then
    MsgBox "Pilih JatuhTempo.."
    Combo33.SetFocus
    Exit Sub
End If

    CR.SelectionFormula = "Totext({Penjualan.JatuhTempo})='" & CDate(Combo33) & "' AND {Penjualan.SISA}<>0"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo33, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap Piutang jatuh tempo.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command19_Click()
    If Combo32 = "" Then
        MsgBox "JatuhTempo awal kosong", , "Informasi"
        Combo32.SetFocus
        Exit Sub
    Else
        If Combo31 < Combo32 Or Combo32 > Combo31 Then
            MsgBox "JatuhTempo terbalik"
            Combo31.SetFocus
            Exit Sub
        ElseIf Combo31 = Combo32 Then
            MsgBox "pilih JatuhTempo yang berbeda"
            Combo31.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{penjualan.JatuhTempo} in date (" & Combo32 & ") to date (" & Combo31 & ") AND {penjualan.SISA}<>0"
    CR.Formulas(0) = "TGLAWAL='" & Format(Combo32, "dd-mmm-yyyy") & "'"
    CR.Formulas(1) = "TGLAKHIR='" & Format(Combo31, "dd-mmm-yyyy") & "'"
    CR.ReportFileName = App.Path & "\Lap piutang jatuh tempo.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command21_Click()
    Call BukaDB
    RSPenjualan.Open "select * from Penjualan where month(JatuhTempo)='" & Val(Left(Combo35, 2)) & "' and year(JatuhTempo)='" & (Combo34) & "' and sisa<>0", Conn
    If RSPenjualan.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo35.SetFocus
    End If
    CR.SelectionFormula = "Month({Penjualan.JatuhTempo})=" & Val(Left(Combo35, 2)) & " and Year({Penjualan.JatuhTempo})=" & Val(Combo34.Text) & " and {Penjualan.sisa}<>0"
    CR.ReportFileName = App.Path & "\Lap Piutang jatuh tempo.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub



