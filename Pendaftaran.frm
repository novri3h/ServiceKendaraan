VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pendaftaran 
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6015
      Begin VB.TextBox TxtTanggal 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4080
         TabIndex        =   15
         Top             =   240
         Width           =   1650
      End
      Begin MSAdodcLib.Adodc ADO 
         Height          =   375
         Left            =   3600
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox TxtNomor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   2000
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   720
         Width           =   2000
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   4500
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1440
         Width           =   4500
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tanggal"
         Height          =   315
         Left            =   3240
         TabIndex        =   16
         Top             =   240
         Width           =   800
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nomor"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No Polisi"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Kendaraan"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1005
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pendaftaran Service"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Pendaftaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
Call BukaDB
ADO.ConnectionString = PathData
ADO.RecordSource = "Pendaftaran"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
TxtTanggal = Date
Call Auto
End Sub

Private Sub Form_Load()
Call Semula
End Sub

Private Sub Auto()
Call BukaDB
RSPendaftaran.Open "select * from Pendaftaran Where Nomor In(Select Max(Nomor)From Pendaftaran)Order By Nomor Desc", Conn
RSPendaftaran.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSPendaftaran
        If .EOF Then
            Urutan = Format(Date, "yymmdd") + "0001"
            TxtNomor = Urutan
        Else
            If Left(!Nomor, 6) <> Format(Date, "yymmdd") Then
                Urutan = Format(Date, "yymmdd") + "0001"
            Else
                Hitung = (!Nomor) + 1
                Urutan = Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
            End If
        End If
        TxtNomor = Urutan
    End With
End Sub

Private Sub CmdBatal_Click()
Call Semula
Text1.SetFocus
End Sub

Private Sub CmdHapus_Click()
If TxtNomor = "" Or TxtTanggal = "" Or Text1 = "" Or Text2 = "" Or Text3 = "" Then
    MsgBox "data yang akan dihapus belum diisi"
    DG.SetFocus
Else
    Pesan = MsgBox("yakin akan dihapus", vbYesNo)
    If Pesan = vbYes Then
        Call BukaDB
        Hapus = "delete * from Pendaftaran where Nomor='" & TxtNomor & "'"
        Conn.Execute Hapus
        Text1.SetFocus
        Call Semula
    Else
        Text1.SetFocus
        Call Semula
    End If
End If
End Sub

Private Sub CmdSimpan_Click()
If TxtNomor = "" Or TxtTanggal = "" Or Text1 = "" Or Text2 = "" Or Text3 = "" Then
    MsgBox "Data belum lengkap"
Else
    Call BukaDB
    RSPendaftaran.Open "select * from Pendaftaran where Nomor='" & TxtNomor & "' and nopol='" & Text1 & "'", Conn
    If RSPendaftaran.EOF Then
        Simpan = "insert into Pendaftaran(Nomor,Tanggal,Nopol,Nama,Kendaraan) values ('" & TxtNomor & "','" & TxtTanggal & "','" & Text1 & "','" & Text2 & "','" & Text3 & "')"
        Conn.Execute Simpan
        Text1.SetFocus
        Call Semula
    Else
        ubah = "update Pendaftaran set Nama='" & Text2 & "',kendaraan='" & Text3 & "' where Nopol='" & Text1 & "' and nomor='" & TxtNomor & "'"
        Conn.Execute ubah
        Text1.SetFocus
        Call Semula
    End If
End If
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

Sub Kosongkan()
Text1 = ""
Text2 = ""
Text3 = ""
End Sub

Sub Tampilkan()
On Error Resume Next
Text2 = RSPendaftaran!nama
Text3 = RSPendaftaran!Kendaraan
End Sub

Sub Semula()
Call Kosongkan
Form_Activate
End Sub

Private Sub DG_Click()
On Error Resume Next
TxtNomor = DG.Columns(0)
TxtTanggal = DG.Columns(1)
Text1 = DG.Columns(2)
Text2 = DG.Columns(3)
Text3 = DG.Columns(4)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Text1.MaxLength = 12
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Call BukaDB
    RSPendaftaran.Open "select * from Pendaftaran where Nopol='" & Text1 & "'", Conn
    If Not RSPendaftaran.EOF Then
        Call Tampilkan
        Text2.SetFocus
    Else
        Text2.SetFocus
    End If
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then CmdSimpan.SetFocus
End Sub

Private Sub TxtCari_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Call BukaDB
    ADO.ConnectionString = PathData
    ADO.RecordSource = "select * from Pendaftaran where Nama like '%" & TxtCari & "%'"
    ADO.Refresh
    If ADO.Recordset.EOF Then
        MsgBox "Nama Pendaftaran tidak ditemukan"
    End If
    ADO.Refresh
End If
End Sub




