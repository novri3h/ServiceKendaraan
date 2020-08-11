VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pembelian 
   Caption         =   "Pembelian"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox HargaBeli 
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
      Left            =   5760
      TabIndex        =   5
      Top             =   720
      Width           =   1250
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
      Left            =   8880
      TabIndex        =   3
      Top             =   150
      Width           =   735
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
      Left            =   7680
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TxtKodeEdit 
      Height          =   350
      Left            =   2640
      TabIndex        =   14
      Text            =   "Kode Barang ?"
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton CmdHapusSemua 
      Caption         =   "Ba&tal Semua"
      Height          =   350
      Left            =   3960
      TabIndex        =   15
      Top             =   4680
      Width           =   1150
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit Jumlah"
      Height          =   350
      Left            =   2640
      TabIndex        =   12
      Top             =   4200
      Width           =   1200
   End
   Begin VB.TextBox TxtKodeBatal 
      Height          =   350
      Left            =   1440
      TabIndex        =   13
      Text            =   "Kode Barang ??"
      Top             =   4680
      Width           =   1200
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
      TabIndex        =   6
      Top             =   720
      Width           =   885
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "Pembelian.frx":0000
      Height          =   2895
      Left            =   120
      TabIndex        =   30
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
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   3285
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
      TabIndex        =   4
      Top             =   720
      Width           =   1250
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1800
      Top             =   5280
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
      TabIndex        =   11
      Top             =   4200
      Width           =   1200
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   240
      TabIndex        =   10
      Top             =   4200
      Width           =   1200
   End
   Begin VB.TextBox TxtDibayar 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   6000
      TabIndex        =   7
      Top             =   4560
      Width           =   1250
   End
   Begin VB.TextBox TxtFaktur 
      Height          =   350
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1250
   End
   Begin VB.TextBox TxtUangMuka 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   8520
      TabIndex        =   9
      Top             =   4560
      Width           =   1250
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   8520
      TabIndex        =   8
      Top             =   4200
      Width           =   1250
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   405
      Left            =   120
      Top             =   5280
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   714
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
   Begin VB.Label LblKodePemasok 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6600
      TabIndex        =   33
      Top             =   120
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9720
      Y1              =   600
      Y2              =   600
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
      TabIndex        =   32
      Top             =   720
      Width           =   1500
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
      TabIndex        =   31
      Top             =   720
      Width           =   3435
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      Height          =   345
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pemasok"
      Height          =   350
      Left            =   2280
      TabIndex        =   28
      Top             =   120
      Width           =   900
   End
   Begin VB.Label LblItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4560
      TabIndex        =   27
      Top             =   4200
      Width           =   600
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item"
      Height          =   345
      Left            =   3960
      TabIndex        =   26
      Top             =   4200
      Width           =   600
   End
   Begin VB.Label LblKembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6000
      TabIndex        =   25
      Top             =   4920
      Width           =   1245
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
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   5160
      TabIndex        =   23
      Top             =   4560
      Width           =   795
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6000
      TabIndex        =   22
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   345
      Left            =   5160
      TabIndex        =   21
      Top             =   4200
      Width           =   795
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Faktur"
      Height          =   350
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   855
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
      TabIndex        =   19
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa"
      Height          =   345
      Left            =   7440
      TabIndex        =   18
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Label LblSisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   8520
      TabIndex        =   17
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tempo"
      Height          =   345
      Left            =   7440
      TabIndex        =   16
      Top             =   4200
      Width           =   1005
   End
End
Attribute VB_Name = "Pembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call BukaDB
ADO.ConnectionString = PathData '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBeli.mdb"
ADO.RecordSource = "Transaksi"
Set DG.DataSource = ADO
DG.Refresh
'Option1.Value = True
'Call Auto
End Sub

Private Sub Form_Load()
Call Tabel_Kosong
Call BukaDB
RSPemasok.Open "pemasok", Conn
Combo1.Clear
Do Until RSPemasok.EOF
    Combo1.AddItem RSPemasok!namapms
    RSPemasok.MoveNext
Loop

Combo2.Enabled = False
TxtUangMuka.Enabled = False
Combo2.Clear
For qq = 1 To 30
    Combo2.AddItem qq
Next qq
'Date = Format(Date, "dd-mm-yyyy")
CmdSimpan.Enabled = False
TxtKodeBatal.Visible = False
TxtKodeEdit.Visible = False
TxtKodeBarang.Enabled = False
TxtJumlah.Enabled = False
'Call Auto
Option1.Value = True
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Combo1 = "" Then
        MsgBox "pilih nama pemasok...!"
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
    RSPemasok.Open "Select * from Pemasok where namapms='" & Combo1 & "'", Conn
    If Not RSPemasok.EOF Then
        LblKodePemasok = RSPemasok!kodepms
        TxtKodeBarang.Enabled = True
        TxtJumlah.Enabled = True
    Else
        MsgBox "Nama pemasok tdak terdaftar"
        Combo1.SetFocus
    End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    TxtDibayar.Enabled = True
    Combo2.Enabled = False
    TxtUangMuka.Enabled = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    TxtDibayar.Enabled = False
    Combo2.Enabled = True
    TxtUangMuka.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
    LblJam = Time$
End Sub


'mencari nomor otomatis
Private Sub Auto()
Call BukaDB
RSPembelian.Open "select * from Pembelian Where Faktur In(Select Max(Faktur)From Pembelian)Order By Faktur Desc", Conn
RSPembelian.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSPembelian
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
ADO.ConnectionString = PathData '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBeli.mdb"
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
LblKodePemasok = ""
TxtKodeBarang = ""
Call Tabel_Kosong
LblItem = ""
LblTotal = ""
TxtDibayar = ""
LblKembali = ""
Combo2 = ""
TxtUangMuka = ""
LblSisa = ""
Combo1.SetFocus
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
'On Error GoTo salah
If Option1.Value = True Then

    If LblKodePemasok = "" Or TxtDibayar = "" Then
        MsgBox "Data belum lengkap"
        Exit Sub
    Else
        If LblItem = "" Then
            MsgBox "tidak ada Transaksi pembelian"
            Exit Sub
        End If
    End If
ElseIf Option2.Value = True Then
    If LblKodePemasok = "" Or Combo2 = "" Or TxtUangMuka = "" Then
        MsgBox "Data belum lengkap"
        Exit Sub
    Else
        If LblItem = "" Then
            MsgBox "tidak Transaksi pembelian"
            Exit Sub
        End If
    End If
End If
    
    Call BukaDB
    If Option1.Value = True Then
        Dim JualTunai As String
        JualTunai = "Insert Into Pembelian(Faktur,Tanggal,jenis,JmlItem,JmlTotal,Dibayar,Kembali,KodeKsr,KodePms)" & _
        "values('" & TxtFaktur & "','" & Date & "','" & Option1.Caption & "','" & LblItem & "','" & LblTotal & "'," & _
        "'" & TxtDibayar & "','" & LblKembali & "','" & Menu.STBar.Panels(1).Text & "','" & LblKodePemasok & "')"
        Conn.Execute (JualTunai)
        
        Dim Kas1 As String
        Kas1 = "insert into kas(tanggal,keterangan,pengeluaran) values " & _
        "('" & Date & "','Pembelian tunai dari " & Combo1 & "','" & LblTotal & "')"
        Conn.Execute Kas1
        
    ElseIf Option2.Value = True Then
        Dim JualKredit As String
        JualKredit = "Insert Into Pembelian(Faktur,Tanggal,jenis,JmlItem,JmlTotal,DP,sisa,tempo,jatuhtempo,KodeKsr,KodePms)" & _
        "values('" & TxtFaktur & "','" & Date & "','" & Option2.Caption & "','" & LblItem & "', " & _
        "'" & LblTotal & "','" & TxtUangMuka & "','" & LblSisa & "','" & Combo2 & "','" & Date + CDate(Combo2) & "','" & Menu.STBar.Panels(1).Text & "','" & LblKodePemasok & "')"
        Conn.Execute (JualKredit)
        
        Dim Kas2 As String
        Kas2 = "insert into kas(tanggal,keterangan,pengeluaran) values " & _
        "('" & Date & "','Pembelian kredit dari " & Combo1 & "','" & TxtUangMuka & "')"
        Conn.Execute Kas2
        
    End If
    
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF
        If ADO.Recordset!kode <> vbNullString Then
            Dim SQLTambahDetail As String
            SQLTambahDetail = "Insert Into DetailBeli(Faktur,Kodebrg,namabarang,hargabeli,JmlBeli,subtotal) " & _
            "values ('" & TxtFaktur & "','" & ADO.Recordset!kode & "','" & ADO.Recordset!nama & "','" & ADO.Recordset!Harga & "','" & ADO.Recordset!jumlah & "','" & ADO.Recordset!Total & "')"
            Conn.Execute (SQLTambahDetail)
        End If
    ADO.Recordset.MoveNext
    Loop
        
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF
        If ADO.Recordset!kode <> vbNullString Then
            Call BukaDB
            RSBarang.Open "Select * from Barang where Kodebrg='" & ADO.Recordset!kode & "'", Conn
            If Not RSBarang.EOF Then
                Dim TambahBarang1 As String
                TambahBarang1 = "update barang set jumlahbrg='" & RSBarang!jumlahbrg + ADO.Recordset!jumlah & "' where kodebrg='" & ADO.Recordset!kode & "'"
                Conn.Execute (TambahBarang1)
            Else
                Dim TambahBarang2 As String
                TambahBarang2 = "Insert Into Barang(Kodebrg,NamaBrg,HargaBeli,HargaJual,JumlahBrg)" & _
                "values('" & ADO.Recordset!kode & "','" & ADO.Recordset!nama & "','" & ADO.Recordset!Harga & "','" & ADO.Recordset!Harga * 1.5 & "','" & ADO.Recordset!jumlah & "')"
                Conn.Execute (TambahBarang2)
            End If
        End If
    ADO.Recordset.MoveNext
    Loop
    
    Bersihkan
    Form_Activate
    TxtFaktur.SetFocus
    TxtFaktur = ""
On Error GoTo 0
Exit Sub
salah:
MsgBox "Nomor faktur ganda, ganti dengan nomor lain"
TxtFaktur.SetFocus
    'Call Cetak
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
    HargaBeli = ""
    Combo1 = ""
    LblKodePemasok = ""
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
        MsgBox "Faktur wajib diisi"
        TxtFaktur.SetFocus
        Exit Sub
    Else
        Call BukaDB
        RSPembelian.Open "select * from pembelian where faktur='" & TxtFaktur & "'", Conn
        If RSPembelian.EOF Then
            Combo1.SetFocus
            Exit Sub
        Else
            MsgBox "nomor faktur ganda, ganti dengan nomor lain"
            TxtFaktur.SetFocus
        End If
    End If
End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub TxtJumlah_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    TxtJumlah = ""
    HargaBeli = ""
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
        Total = HargaBeli * Val(TxtJumlah)
        Dim Simpan As String
        Simpan = "insert into Transaksi(kode,nama,harga,jumlah,total) values " & _
        "('" & TxtKodeBarang & "','" & NamaBarang & "','" & HargaBeli & "','" & TxtJumlah & "','" & Total & "')"
        Conn.Execute Simpan
        Form_Activate
        ADO.Refresh
        DG.Refresh
        LblTotal = Format(TotalHarga, "###,###,###")
        LblItem = Format(TotalItem, "#,###,###")
        Call Lagi
    End If
End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub TxtKodeBarang_KeyPress(KeyAscii As Integer)
'TxtKodeBarang.MaxLength = 6
If KeyAscii = 27 Then Unload Me
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Call BukaDB
    RSBarang.Open "select * from barang where kodebrg='" & TxtKodeBarang & "'", Conn
    RSBarang.Requery
    If Not RSBarang.EOF Then
        NamaBarang = RSBarang!namabrg
        HargaBeli = RSBarang!HargaBeli
        HargaBeli.SetFocus
        'TxtJumlah.SetFocus
        'TxtJumlah = 1
        Exit Sub
    Else
        DaftarBarangBeli.Show
    End If
End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub hargabeli_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then TxtJumlah.SetFocus
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub


Private Sub TxtKodeBatal_KeyPress(KeyAscii As Integer)
'TxtKodeBatal.MaxLength = 6
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
'TxtKodeEdit.MaxLength = 6
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
HargaBeli = ""
TxtJumlah = ""
Total = ""
TxtKodeBarang.SetFocus
End Sub
