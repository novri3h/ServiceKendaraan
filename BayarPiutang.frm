VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TerimaPiutang 
   Caption         =   "Penerimaan Piutang"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
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
   ScaleHeight     =   4035
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   1680
      Width           =   1800
   End
   Begin VB.TextBox TxtDibayar 
      Height          =   350
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   1800
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   1800
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1800
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1680
      Width           =   1800
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   375
      Left            =   5760
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "ADO"
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
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "BayarPiutang.frx":0000
      Height          =   1800
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3175
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "NomorBayar"
         Caption         =   "NomorBayar"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "TanggalTerima"
         Caption         =   "TanggalTerima"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Faktur"
         Caption         =   "Faktur"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Dibayar"
         Caption         =   "Dibayar"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Sisa"
         Caption         =   "Sisa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Keterangan"
         Caption         =   "Keterangan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1140,095
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker TanggalTerima 
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   91881473
      CurrentDate     =   40568
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa"
      Height          =   345
      Left            =   3960
      TabIndex        =   18
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   3960
      TabIndex        =   17
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah Piutang"
      Height          =   345
      Left            =   3960
      TabIndex        =   16
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Pelanggan"
      Height          =   345
      Left            =   3960
      TabIndex        =   15
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal Jual"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Faktur"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal Terima"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Pembayaran"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label LblSisa 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5880
      TabIndex        =   10
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label LblJumlahPiutang 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5880
      TabIndex        =   9
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label LblNamaPelanggan 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5880
      TabIndex        =   8
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label LblTanggalJual 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label LblNomorBayar 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "TerimaPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdTutup_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call BukaDB
ADO.ConnectionString = PathData
ADO.RecordSource = "TerimaPiutang"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
Call Auto
RSPenjualan.Open "select * from Penjualan where jenis='kredit' and sisa<>0", Conn
Combo1.Clear
Do Until RSPenjualan.EOF
    Combo1.AddItem RSPenjualan!faktur
    RSPenjualan.MoveNext
Loop
TanggalTerima.Value = Date
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Combo1 = "" Then
        MsgBox "pilih Nomor Faktur...!"
        Combo1.SetFocus
    Else
        TxtDibayar.SetFocus
    End If
End If
End Sub

Private Sub Combo1_Click()
    Call BukaDB
    RSPenjualan.Open "Select Penjualan.tanggal,sisa,Pelanggan.namaPlg from Penjualan,Pelanggan where faktur='" & Combo1 & "' and Penjualan.kodePlg=Pelanggan.kodePlg", Conn
    If Not RSPenjualan.EOF Then
        LblTanggalJual = RSPenjualan!tanggal
        LblNamaPelanggan = RSPenjualan!NamaPlg
        'LblKodePelanggan = RSPenjualan!kodePlg
        LblJumlahPiutang = Format(RSPenjualan!sisa, "###,###,###")
    Else
        MsgBox "Nama Penjualan tdak terdaftar"
        Combo1.SetFocus
    End If
    
End Sub

'mencari nomor otomatis
Private Sub Auto()
Call BukaDB
RSPiutang.Open "select * from TerimaPiutang Where NomorBayar In(Select Max(NomorBayar)From TerimaPiutang)Order By NomorBayar Desc", Conn
RSPiutang.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSPiutang
        If .EOF Then
            Urutan = "PI" + Format(Date, "yymmdd") + "01"
            LblNomorBayar = Urutan
        Else
            If Mid(!NomorBayar, 3, 6) <> Format(Date, "yymmdd") Then
                Urutan = "PI" + Format(Date, "yymmdd") + "01"
            Else
                Hitung = Right(!NomorBayar, 2) + 1
                Urutan = "PI" + Format(Date, "yymmdd") + Right("00" & Hitung, 2)
            End If
        End If
        LblNomorBayar = Urutan
    End With
End Sub

Private Sub Bersihkan()
Combo1 = ""
LblTanggalJual = ""
LblNamaPelanggan = ""
LblJumlahPiutang = ""
TxtDibayar = ""
LblSisa = ""
TanggalTerima.SetFocus
End Sub

Private Sub TxtDibayar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtDibayar = "" Then
            MsgBox "Pembayaran belum diisi"
            TxtDibayar.SetFocus
            Exit Sub
        Else
            TxtDibayar = Format(TxtDibayar, "###,###,###")
            If TxtDibayar = LblJumlahPiutang Then
                LblSisa = TxtDibayar - LblJumlahPiutang
            Else
                LblSisa = Format(LblJumlahPiutang - TxtDibayar, "###,###,###")
            End If
        End If
        CmdSimpan.SetFocus
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub CmdSimpan_Keypress(KeyAscii As Integer)
    If KeyAscii = 27 Then Call Bersihkan
End Sub

Private Sub CmdSimpan_Click()
If Combo1 = "" Or TxtDibayar = "" Then
    MsgBox "Data belum lengkap"
    Exit Sub
Else
    Call BukaDB
    
    Dim SimpanBayar As String
    SimpanBayar = "Insert Into TerimaPiutang(NomorBayar,TanggalTerima,FAKTUR,Dibayar,sisa,KETERANGAN)" & _
    "values('" & LblNomorBayar & "','" & TanggalTerima & "','" & Combo1 & "','" & TxtDibayar & "','" & LblSisa & "','PIUTANG')"
    Conn.Execute SimpanBayar
    
    Dim lUNASI As String
    lUNASI = "update TerimaPiutang set keterangan='LUNAS' where SISA=0 AND NOMORBAYAR='" & LblNomorBayar & "' AND FAKTUR='" & Combo1 & "'"
    Conn.Execute lUNASI
    
    Dim UbahJual As String
    UbahJual = "update Penjualan set sisa='" & LblSisa & "' where FAKTUR='" & Combo1 & "'"
    Conn.Execute UbahJual
    
    
    Dim Kas As String
    Kas = "insert into kas(Tanggal,keterangan,pemasukan) values " & _
    "('" & Date & "','" & TerimaPiutang.Caption + Space(1) + LblNamaPelanggan & "','" & TxtDibayar & "')"
    Conn.Execute Kas
    
    Bersihkan
    Form_Activate
End If
End Sub


Private Sub CmdBatal_Click()
Call Bersihkan
End Sub





