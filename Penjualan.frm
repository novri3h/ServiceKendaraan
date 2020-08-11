VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Penjualan 
   Caption         =   "Penjualan"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
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
   ScaleHeight     =   5430
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   8520
      TabIndex        =   6
      Top             =   4200
      Width           =   1250
   End
   Begin VB.TextBox TxtUangMuka 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   8520
      TabIndex        =   7
      Top             =   4560
      Width           =   1250
   End
   Begin VB.TextBox TxtFaktur 
      Height          =   350
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   1250
   End
   Begin VB.TextBox TxtDibayar 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   6000
      TabIndex        =   5
      Top             =   4560
      Width           =   1250
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   1200
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batalkan Item"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1440
      TabIndex        =   10
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1800
      Top             =   5040
   End
   Begin VB.TextBox TxtKodeBarang 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   1250
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   3465
   End
   Begin VB.TextBox TxtJumlah 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7200
      TabIndex        =   1
      Top             =   720
      Width           =   885
   End
   Begin VB.TextBox TxtKodeBatal 
      Height          =   350
      Left            =   1440
      TabIndex        =   12
      Text            =   "Kode Barang ??"
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit Jumlah"
      Height          =   350
      Left            =   2640
      TabIndex        =   11
      Top             =   4200
      Width           =   1200
   End
   Begin VB.CommandButton CmdHapusSemua 
      Caption         =   "Ba&tal Semua"
      Height          =   350
      Left            =   3960
      TabIndex        =   14
      Top             =   4680
      Width           =   1150
   End
   Begin VB.TextBox TxtKodeEdit 
      Height          =   350
      Left            =   2640
      TabIndex        =   13
      Text            =   "Kode Barang ?"
      Top             =   4680
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tunai"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7650
      TabIndex        =   2
      Top             =   165
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Kredit"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8685
      TabIndex        =   3
      Top             =   195
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "Penjualan.frx":0000
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   0   'False
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Kode"
         Caption         =   "Kode"
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
         DataField       =   "Nama"
         Caption         =   "Nama"
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
         DataField       =   "Harga"
         Caption         =   "Harga"
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
         DataField       =   "Jumlah"
         Caption         =   "Jumlah"
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
         DataField       =   "Total"
         Caption         =   "Total"
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
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4004,788
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1500,095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   350
      Left            =   120
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
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
   Begin VB.Label LblKodePelanggan 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6720
      TabIndex        =   33
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tempo"
      Height          =   345
      Left            =   7440
      TabIndex        =   32
      Top             =   4200
      Width           =   1005
   End
   Begin VB.Label LblSisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   8520
      TabIndex        =   31
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa"
      Height          =   345
      Left            =   7440
      TabIndex        =   30
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Uang Muka"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7440
      TabIndex        =   29
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Faktur"
      Height          =   345
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   345
      Left            =   5160
      TabIndex        =   27
      Top             =   4200
      Width           =   795
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6000
      TabIndex        =   26
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   5160
      TabIndex        =   25
      Top             =   4560
      Width           =   795
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   345
      Left            =   5160
      TabIndex        =   24
      Top             =   4920
      Width           =   795
   End
   Begin VB.Label LblKembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6000
      TabIndex        =   23
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item"
      Height          =   345
      Left            =   3960
      TabIndex        =   22
      Top             =   4200
      Width           =   600
   End
   Begin VB.Label LblItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4560
      TabIndex        =   21
      Top             =   4200
      Width           =   600
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pelanggan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   20
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      Height          =   345
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   855
   End
   Begin VB.Label NamaBarang 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   18
      Top             =   720
      Width           =   3435
   End
   Begin VB.Label HargaJual 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      TabIndex        =   17
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label Total 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8160
      TabIndex        =   16
      Top             =   720
      Width           =   1500
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9720
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "Penjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call BukaDB
ADO.ConnectionString = PathData
ADO.RecordSource = "Transaksi"
Set DG.DataSource = ADO
DG.Refresh
Call Auto
CmdSimpan.Enabled = False
End Sub

Private Sub Form_Load()
Call Tabel_Kosong
Call BukaDB
RSPelanggan.Open "Pelanggan", Conn
Combo1.Clear
Do Until RSPelanggan.EOF
    Combo1.AddItem RSPelanggan!NamaPlg
    RSPelanggan.MoveNext
Loop

Combo2.Enabled = False
TxtUangMuka.Enabled = False
Combo2.Clear
For qq = 1 To 30
    Combo2.AddItem qq
Next qq
TxtKodeBatal.Visible = False
TxtKodeEdit.Visible = False
TxtFaktur.Enabled = False
Option1.Value = True
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Combo1 = "" Then
        MsgBox "pilih nama Pelanggan...!"
        Combo1.SetFocus
    Else
        TxtKodeBarang.Enabled = True
        TxtKodeBarang.SetFocus
    End If
End If
If KeyAscii = 27 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Combo1_Click()
    Call BukaDB
    RSPelanggan.Open "Select * from pelanggan where namaPlg='" & Combo1 & "'", Conn
    If Not RSPelanggan.EOF Then
        LblKodePelanggan = RSPelanggan!kodePlg
        LblKodePelanggan.Enabled = False
        TxtKodeBarang.Enabled = True
    Else
        MsgBox "Nama Pelanggan tdak terdaftar"
        Combo1.SetFocus
    End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    Combo1.Enabled = False
    Combo1 = ""
    Combo1 = "PELANGGAN UMUM"
    LblKodePelanggan = "UMUM"
    TxtDibayar.Enabled = True
    Combo2.Enabled = False
    TxtUangMuka.Enabled = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    Combo1.Enabled = True
    TxtDibayar.Enabled = False
    TxtDibayar = ""
    LblKembali = ""
    Combo2.Enabled = True
    TxtUangMuka.Enabled = True
    Combo1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
    LblJam = Time$
End Sub


'mencari nomor otomatis
Private Sub Auto()
Call BukaDB
RSPenjualan.Open "select * from Penjualan Where Faktur In(Select Max(Faktur)From Penjualan)Order By Faktur Desc", Conn
RSPenjualan.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSPenjualan
        If .EOF Then
            Urutan = Format(Date, "yymmdd") + "0001"
            TxtFaktur = Urutan
        Else
            If Left(!faktur, 6) <> Format(Date, "yymmdd") Then
                Urutan = Format(Date, "yymmdd") + "0001"
            Else
                Hitung = (!faktur) + 1
                Urutan = Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
            End If
        End If
        TxtFaktur = Urutan
    End With
End Sub

Function Tabel_Kosong()
'On Error Resume Next
Call BukaDB
ADO.ConnectionString = PathData '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\adoJual.mdb"
ADO.RecordSource = "Transaksi"
ADO.Refresh
If ADO.Recordset.RecordCount <> 0 Then
    Do While Not ADO.Recordset.EOF
        ADO.Recordset.Delete
        ADO.Recordset.MoveNext
    Loop
End If
End Function

Private Sub Bersihkan()
Combo1 = ""
LblKodePelanggan = ""
TxtKodeBarang = ""
Call Tabel_Kosong
LblItem = ""
LblTotal = ""
TxtDibayar = ""
LblKembali = ""
Combo2 = ""
TxtUangMuka = ""
LblSisa = ""
TxtKodeBarang.SetFocus
End Sub

Private Sub TxtDibayar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtDibayar = "" Or Val(TxtDibayar) < (LblTotal) Then
            MsgBox "Jumlah Pembayaran Kurang"
            TxtDibayar.SetFocus
        Else
            TxtDibayar = Format(TxtDibayar, "###,###,###")
            If TxtDibayar = LblTotal Then
                LblKembali = TxtDibayar - LblTotal
            Else
                LblKembali = Format(TxtDibayar - LblTotal, "###,###,###")
            End If
        CmdSimpan.Enabled = True
        CmdSimpan.SetFocus
        End If
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub CmdSimpan_Keypress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        CmdSimpan.Enabled = False
        TxtDibayar = ""
        TxtUangMuka = ""
        TxtDibayar.SetFocus
    End If
End Sub

Private Sub CmdSimpan_Click()
If Option1.Value = True Then

    If TxtDibayar = "" Then
        MsgBox "Data belum lengkap"
        Exit Sub
    Else
        If LblItem = "" Then
            MsgBox "tidak ada Transaksi Penjualan"
            Exit Sub
        End If
    End If
ElseIf Option2.Value = True Then
    If LblKodePelanggan = "" Or Combo2 = "" Or TxtUangMuka = "" Or LblSisa = "" Then
        MsgBox "Data belum lengkap"
        Exit Sub
    Else
        If LblItem = "" Then
            MsgBox "tidak ada transaksi Penjualan"
            Exit Sub
        End If
    End If
End If
    
    Call BukaDB
    If Option1.Value = True Then
        Dim JualTunai As String
        JualTunai = "Insert Into Penjualan(Faktur,Tanggal,jenis,JmlItem,JmlTotal,Dibayar,Kembali,KodeKsr,KodePlg)" & _
        "values('" & TxtFaktur & "','" & Date & "','" & Option1.Caption & "','" & LblItem & "','" & LblTotal & "'," & _
        "'" & TxtDibayar & "','" & LblKembali & "','" & Menu.STBar.Panels(1).Text & "','" & LblKodePelanggan & "')"
        Conn.Execute (JualTunai)
        
        Dim Kas1 As String
        Kas1 = "insert into kas(tanggal,keterangan,pemasukan) values " & _
        "('" & Date & "','Penjualan tunai','" & LblTotal & "')"
        Conn.Execute Kas1

    ElseIf Option2.Value = True Then
        Dim JualKredit As String
        JualKredit = "Insert Into Penjualan(Faktur,Tanggal,jenis,JmlItem,JmlTotal,DP,sisa,tempo,jatuhtempo,KodeKsr,KodePlg)" & _
        "values('" & TxtFaktur & "','" & Date & "','" & Option2.Caption & "','" & LblItem & "', " & _
        "'" & LblTotal & "','" & TxtUangMuka & "','" & LblSisa & "','" & Combo2 & "','" & Date + CDate(Combo2) & "','" & Menu.STBar.Panels(1).Text & "','" & LblKodePelanggan & "')"
        Conn.Execute (JualKredit)
        
        Dim Kas2 As String
        Kas2 = "insert into kas(tanggal,keterangan,pemasukan) values " & _
        "('" & Date & "','Penjualan kredit ke " & Combo1 & "','" & TxtUangMuka & "')"
        Conn.Execute Kas2

    End If
    
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF
        If ADO.Recordset!kode <> vbNullString Then
            Dim SQLTambahDetail As String
            SQLTambahDetail = "Insert Into DetailJual(Faktur,Kodebrg,namabarang,hargaJual,JmlJual,subtotal) " & _
            "values ('" & TxtFaktur & "','" & ADO.Recordset!kode & "','" & ADO.Recordset!nama & "','" & ADO.Recordset!Harga & "','" & ADO.Recordset!jumlah & "','" & ADO.Recordset!Total & "')"
            Conn.Execute (SQLTambahDetail)
        End If
    ADO.Recordset.MoveNext
    Loop
        
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF
        'If ADO.Recordset!kode <> vbNullString Then
            Call BukaDB
            RSBarang.Open "Select * from Barang where Kodebrg='" & ADO.Recordset!kode & "'", Conn
            If Not RSBarang.EOF Then
                Dim KurangiStok As String
                KurangiStok = "update barang set jumlahbrg='" & RSBarang!jumlahbrg - ADO.Recordset!jumlah & "' where kodebrg='" & ADO.Recordset!kode & "'"
                Conn.Execute (KurangiStok)
            End If
        'End If
    ADO.Recordset.MoveNext
    Loop
    
    Bersihkan
    Form_Activate
    Call Cetak
End Sub

Private Sub CmdSimpan_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyAscii = vbKeyF1 Then

End Sub

Private Sub CmdBatal_Click()
TxtKodeEdit.Visible = False
TxtKodeBatal.Visible = True
TxtKodeBatal.SetFocus
End Sub

Private Sub CmdEdit_Click()
TxtKodeBatal.Visible = False
TxtKodeEdit.Visible = True
TxtKodeEdit.SetFocus
End Sub

Private Sub CmdHapusSemua_Click()
    NamaBarang = ""
    HargaJual = ""
    TxtJumlah = ""
    Combo1 = ""
    LblKodePelanggan = ""
    TxtDibayar = ""
    LblTotal = ""
    LblItem = ""
    LblKembali = ""
    Combo2 = ""
    TxtUangMuka = ""
    LblKembali = ""
    TxtKodeBarang = ""
    TxtKodeBatal.Visible = False
    TxtKodeEdit.Visible = False
    Call Tabel_Kosong
    Call Auto
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

'mencari total harga
Function TotalHarga()
    Dim RS1 As New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    RS1.Open "select sum(Total) as JumTotal from Transaksi", Conn
    TotalHarga = RS1!JumTotal
End Function

'mencari total item
Function TotalItem()
    Dim RS2 As New ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    RS2.Open "select sum(Jumlah) as JumItem from Transaksi", Conn
    TotalItem = RS2!Jumitem
End Function

Private Sub TxtFaktur_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TxtFaktur = "" Then
        MsgBox "Fantur wajib diisi"
        TxtFaktur.SetFocus
        Exit Sub
    Else
        Combo1.SetFocus
    End If
End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub TxtJumlah_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    TxtJumlah = ""
    HargaJual = ""
    NamaBarang = ""
    TxtKodeBarang = ""
    TxtKodeBarang.SetFocus
End If

If KeyAscii = 13 Then
    If TxtJumlah = "" Then
        MsgBox "jumlah wajib diisi"
        TxtJumlah.SetFocus
        Exit Sub
    Else
        Call BukaDB
        RSBarang.Open "select * from barang where kodebrg='" & TxtKodeBarang & "'", Conn
        If RSBarang!jumlahbrg < Val(TxtJumlah) Then
            MsgBox "stok hanya ada " & RSBarang!jumlahbrg & ""
            TxtJumlah.SetFocus
            Exit Sub
        Else
            Total = HargaJual * Val(TxtJumlah)
            Dim Simpan As String
            Simpan = "insert into Transaksi(kode,nama,harga,jumlah,total) values " & _
            "('" & TxtKodeBarang & "','" & NamaBarang & "','" & HargaJual & "','" & TxtJumlah & "','" & Total & "')"
            Conn.Execute Simpan
            Form_Activate
            ADO.Refresh
            DG.Refresh
            LblTotal = Format(TotalHarga, "###,###,###")
            LblItem = Format(TotalItem, "#,###,###")
            Call Lagi
        End If
    End If
End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub TxtKodeBarang_KeyPress(KeyAscii As Integer)
TxtKodeBarang.MaxLength = 13
If KeyAscii = 27 Then Unload Me
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Call BukaDB
    RSBarang.Open "select * from barang where kodebrg='" & TxtKodeBarang & "'", Conn
    RSBarang.Requery
    If Not RSBarang.EOF Then
        NamaBarang = RSBarang!namabrg
        HargaJual = RSBarang!HargaJual
        TxtJumlah.SetFocus
        'TxtJumlah = 1
        Exit Sub
    Else
        DaftarBarangJual.Show
    End If
End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub TxtKodeBatal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    TxtKodeBatal.Visible = False
    TxtKodeBatal = ""
    TxtKodeBarang.SetFocus
End If
If KeyAscii = 13 Then
    If TxtKodeBatal = "" Then
        MsgBox "Kode barang wajib diisi"
        TxtKodeBatal.SetFocus
        Exit Sub
    Else
        Call BukaDB
        RSTransaksi.Open "select * from Transaksi where kode='" & TxtKodeBatal & "'", Conn
        If RSTransaksi.EOF Then
            MsgBox "Kode " & TxtKodeBatal & " tidak ada dalam transaksi (ESC = Tutup Kode)"
            TxtKodeBatal.SetFocus
            Exit Sub
        Else
            Dim Hapus As String
            Hapus = "delete * from Transaksi where kode='" & TxtKodeBatal & "'"
            Conn.Execute Hapus
            Form_Activate
            ADO.Refresh
            DG.Refresh
            LblTotal = Format(TotalHarga, "###,###,###")
            LblItem = Format(TotalItem, "#,###,###")
            TxtKodeBatal.Visible = False
            TxtKodeBatal = ""
            Call Lagi
        End If
    End If
End If
End Sub

Private Sub TxtKodeEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    TxtKodeEdit.Visible = False
    TxtKodeEdit = ""
    TxtKodeBarang.SetFocus
End If
If KeyAscii = 13 Then
    If TxtKodeEdit = "" Then
        MsgBox "kode barang yang akan diedit jumlahnya wajib diisi"
        TxtKodeEdit.SetFocus
        Exit Sub
    Else
        Call BukaDB
        RSTransaksi.Open "select * from Transaksi where kode='" & TxtKodeEdit & "'", Conn
        If RSTransaksi.EOF Then
            MsgBox "Kode barang tidak ditemukan dalam transaksi"
            TxtKodeEdit.SetFocus
        Else
            Dim GantiJumlah As Integer
            GantiJumlah = InputBox("ketik jumlah barang pengganti")
            Dim edit As String
            edit = "update Transaksi set jumlah='" & GantiJumlah & "',total='" & RSTransaksi!Harga * GantiJumlah & "' Where kode='" & TxtKodeEdit & "'"
            Conn.Execute edit
            Form_Activate
            ADO.Refresh
            DG.Refresh
            LblTotal = Format(TotalHarga, "###,###,###")
            LblItem = Format(TotalItem, "#,###,###")
            TxtKodeEdit.Visible = False
            TxtKodeEdit = ""
            Call Lagi
        End If
    End If
End If
End Sub

Private Sub TxtUangMuka_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtUangMuka = "" Then
            TxtUangMuka = 0
            CmdSimpan.Enabled = True
            CmdSimpan.SetFocus
        Else
            TxtUangMuka = Format(TxtUangMuka, "###,###,###")
            If TxtUangMuka = LblTotal Then
                LblSisa = LblTotal - TxtUangMuka
            Else
                LblSisa = Format(LblTotal - TxtUangMuka, "###,###,###")
            End If
        CmdSimpan.Enabled = True
        CmdSimpan.SetFocus
        End If
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Sub Lagi()
TxtKodeBarang = ""
NamaBarang = ""
HargaJual = ""
TxtJumlah = ""
Total = ""
TxtKodeBarang.SetFocus
End Sub


'=========================================
Function Cetak()
Tampilkan.Show
Call BukaDB
'cari faktur terakhir
RSPenjualan.Open "select * from Penjualan Where Faktur In(Select Max(Faktur)From Penjualan)Order By Faktur Desc", Conn
Dim JmlHarga, JmlJual, JmlHasil As Double
Dim MGrs As String
Tampilkan.Font = "Courier New"
Tampilkan.Print
Tampilkan.Print
RSKasir.Open "select * From Kasir where KodeKsr= '" & RSPenjualan!kodeksr & "'", Conn
If RSPenjualan!jenis = "Tunai" Then
    Tampilkan.Print Tab(5); "Faktur     :   "; RSPenjualan!faktur
    Tampilkan.Print Tab(5); "Tanggal    :   "; Format(RSPenjualan!tanggal, "DD-MMMM-YYYY")
    Tampilkan.Print Tab(5); "Kasir      :   "; RSKasir!NamaKsr
ElseIf RSPenjualan!jenis = "Kredit" Then
    RSPelanggan.Open "select * from pelanggan where KodePlg= '" & RSPenjualan!kodePlg & "'", Conn
    Tampilkan.Print Tab(5); "Faktur     :   "; RSPenjualan!faktur
    Tampilkan.Print Tab(5); "Tanggal    :   "; Format(RSPenjualan!tanggal, "DD-MMMM-YYYY")
    Tampilkan.Print Tab(5); "Kasir      :   "; RSKasir!NamaKsr
    Tampilkan.Print Tab(5); "Jenis      :   "; RSPenjualan!jenis
    Tampilkan.Print Tab(5); "Pelanggan  :   "; RSPelanggan!NamaPlg
    Tampilkan.Print Tab(5); "Telepon    :   "; RSPelanggan!teleponPlg
End If
MGrs = String$(33, "-")
Tampilkan.Print Tab(5); MGrs

'cari data di tabel detailJual yang fakturnya =di tbl Penjualan
RSDetailJual.Open "select * from DetailJual Where Faktur='" & RSPenjualan!faktur & "'", Conn
RSDetailJual.MoveFirst

No = 0
Do While Not RSDetailJual.EOF
    No = No + 1
    
    Set RSBarang = New ADODB.Recordset
    'cari barang yang kodenya disimpan di tabel detailJual
    RSBarang.Open "select * From Barang where Kodebrg= '" & RSDetailJual!KodeBrg & "'", Conn
    RSBarang.Requery
    Harga = RSBarang!HargaJual
    jumlah = RSDetailJual!JmlJual
    Hasil = Harga * jumlah
    'Tampilkan berulang-ulang kode,nama,harga,jumlah dan total
    Tampilkan.Print Tab(5); No; Space(2); RSBarang!namabrg
    Tampilkan.Print Tab(10); RKanan(jumlah, "##"); Space(1); "X";
    Tampilkan.Print Tab(15); Format(Harga, "###,###,###");
    Tampilkan.Print Tab(25); RKanan(Hasil, "###,###,###")
    RSDetailJual.MoveNext
Loop

'Tampilkan total harga
If RSPenjualan!jenis = "Tunai" Then
    Tampilkan.Print Tab(5); MGrs
    Tampilkan.Print Tab(5); "Total      :";
    Tampilkan.Print Tab(25); RKanan(RSPenjualan!jmlTotal, "###,###,###");
    Tampilkan.Print Tab(5); "Dibayar    :";
    'Tampilkan dibayar
    Tampilkan.Print Tab(25); RKanan(RSPenjualan!dibayar, "###,###,###");
    Tampilkan.Print Tab(5); MGrs
    Tampilkan.Print Tab(5); "Kembali    :";
    'Tampilkan kembalian
    If RSPenjualan!dibayar = RSPenjualan!jmlTotal Then
        Tampilkan.Print Tab(34); RSPenjualan!dibayar - RSPenjualan!jmlTotal
    Else
        Tampilkan.Print Tab(25); RKanan(RSPenjualan!dibayar - RSPenjualan!jmlTotal, "###,###,###");
    End If
    Tampilkan.Print Tab(5); MGrs
ElseIf RSPenjualan!jenis = "Kredit" Then
    Tampilkan.Print Tab(5); MGrs
    Tampilkan.Print Tab(5); "Total      :";
    Tampilkan.Print Tab(25); RKanan(RSPenjualan!jmlTotal, "###,###,###");
    Tampilkan.Print Tab(5); "Tempo      :";
    Tampilkan.Print Tab(25); RKanan(RSPenjualan!tempo, "###,###,###");
    Tampilkan.Print Tab(5); "Jatuh Tempo:";
    Tampilkan.Print Tab(25); RKanan(RSPenjualan!jatuhtempo, "dd-mmm-yyyy");
    Tampilkan.Print Tab(5); "Uang Muka  :";
    Tampilkan.Print Tab(25); RKanan(RSPenjualan!DP, "###,###,###");
    Tampilkan.Print Tab(5); "Sisa       :";
    Tampilkan.Print Tab(25); RKanan(RSPenjualan!sisa, "###,###,###");
    Tampilkan.Print Tab(5); MGrs
End If
Tampilkan.Print
Tampilkan.Print
Tampilkan.Print
'Tampilkan.EndDoc
Conn.Close
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function


