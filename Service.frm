VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Service 
   Caption         =   "Transaksi Service"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
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
   ScaleHeight     =   8175
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNoPol 
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
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1250
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   8760
      TabIndex        =   37
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatalJasa 
      Caption         =   "Batal"
      Height          =   350
      Left            =   5760
      TabIndex        =   34
      Top             =   3600
      Width           =   1000
   End
   Begin MSDataGridLib.DataGrid DG1 
      Bindings        =   "Service.frx":0000
      Height          =   1935
      Left            =   120
      TabIndex        =   31
      Top             =   1560
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Kode"
         Caption         =   "Kode"
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
         DataField       =   "Nama_Jasa"
         Caption         =   "Nama_Jasa"
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
      BeginProperty Column02 
         DataField       =   "Harga"
         Caption         =   "Harga"
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
            Alignment       =   2
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5999,812
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1755,213
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtKodeJasa 
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
      TabIndex        =   2
      Top             =   1080
      Width           =   1250
   End
   Begin VB.TextBox TxtDiskon 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   350
      Left            =   6000
      TabIndex        =   12
      Text            =   "0"
      Top             =   6960
      Width           =   1250
   End
   Begin VB.ComboBox CBODiskon 
      Height          =   345
      ItemData        =   "Service.frx":0013
      Left            =   4680
      List            =   "Service.frx":0015
      TabIndex        =   6
      Text            =   "Discount"
      Top             =   6960
      Width           =   1250
   End
   Begin VB.TextBox HargaJual 
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
      Left            =   5880
      TabIndex        =   4
      Top             =   4080
      Width           =   1250
   End
   Begin VB.TextBox TxtFaktur 
      Height          =   350
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1250
   End
   Begin VB.TextBox TxtDibayar 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   350
      Left            =   8640
      TabIndex        =   7
      Top             =   6600
      Width           =   1250
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   240
      TabIndex        =   8
      Top             =   6600
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
      TabIndex        =   9
      Top             =   6600
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1800
      Top             =   7440
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
      TabIndex        =   3
      Top             =   4080
      Width           =   1250
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
      TabIndex        =   5
      Top             =   4080
      Width           =   885
   End
   Begin VB.TextBox TxtKodeBatal 
      Height          =   350
      Left            =   1440
      TabIndex        =   13
      Text            =   "Kode Barang ??"
      Top             =   7080
      Width           =   1200
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit Jumlah"
      Height          =   350
      Left            =   2640
      TabIndex        =   10
      Top             =   6600
      Width           =   1200
   End
   Begin VB.CommandButton CmdHapusSemua 
      Caption         =   "Ba&tal Semua"
      Height          =   350
      Left            =   240
      TabIndex        =   11
      Top             =   6960
      Width           =   1150
   End
   Begin VB.TextBox TxtKodeEdit 
      Height          =   350
      Left            =   2640
      TabIndex        =   14
      Text            =   "Kode Barang ?"
      Top             =   7080
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DG2 
      Bindings        =   "Service.frx":0017
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3413
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
   Begin MSAdodcLib.Adodc ADO2 
      Height          =   345
      Left            =   120
      Top             =   7440
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
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   345
      Left            =   120
      Top             =   3600
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
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\App Fix\Program Service Kendaraan\DBRetail.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\App Fix\Program Service Kendaraan\DBRetail.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TempJasa"
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
   Begin VB.Label LblNamaPelanggan 
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
      Left            =   4800
      TabIndex        =   38
      Top             =   120
      Width           =   3435
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Biaya"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4680
      TabIndex        =   36
      Top             =   7320
      Width           =   2600
   End
   Begin VB.Label TotalBiaya 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   750
      Left            =   7320
      TabIndex        =   35
      Top             =   7320
      Width           =   2600
   End
   Begin VB.Label LblBiayaJasa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   8160
      TabIndex        =   33
      Top             =   3600
      Width           =   1485
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Biaya Jasa"
      Height          =   345
      Left            =   6840
      TabIndex        =   32
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Label Harga 
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
      TabIndex        =   30
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label NamaJasa 
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
      TabIndex        =   29
      Top             =   1080
      Width           =   5835
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jasa"
      Height          =   345
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Faktur"
      Height          =   345
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Biaya Barang"
      Height          =   345
      Left            =   4680
      TabIndex        =   26
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label LblBiayaBarang 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   6000
      TabIndex        =   25
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   7320
      TabIndex        =   24
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   345
      Left            =   7320
      TabIndex        =   23
      Top             =   6960
      Width           =   1245
   End
   Begin VB.Label LblKembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   8640
      TabIndex        =   22
      Top             =   6960
      Width           =   1245
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item"
      Height          =   345
      Left            =   3960
      TabIndex        =   21
      Top             =   6600
      Width           =   600
   End
   Begin VB.Label LblItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3960
      TabIndex        =   20
      Top             =   7080
      Width           =   600
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Polisi "
      Height          =   345
      Left            =   2280
      TabIndex        =   19
      Top             =   120
      Width           =   1100
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Barang"
      Height          =   345
      Left            =   120
      TabIndex        =   18
      Top             =   4080
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
      TabIndex        =   17
      Top             =   4080
      Width           =   3435
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
      Top             =   4080
      Width           =   1500
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9720
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "Service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call BukaDB

ADO1.ConnectionString = PathData
ADO1.RecordSource = "TempJasa"
Set DG1.DataSource = ADO1
DG1.Refresh

ADO2.ConnectionString = PathData
ADO2.RecordSource = "Transaksi"
Set DG2.DataSource = ADO2
DG2.Refresh

Call Auto
CmdSimpan.Enabled = False
End Sub

Private Sub Form_Load()
Call KosongkanJasa
Call KosongkanBiaya
Call BukaDB

TxtKodeBatal.Visible = False
TxtKodeEdit.Visible = False
TxtFaktur.Enabled = False

For i = 5 To 75 Step 5
    CBODiskon.AddItem i
Next i
TxtNoPol = ""
LblNamaPelanggan = ""
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
    RSPelanggan1.Open "Select * from Pelanggan1 where namaPlg='" & Combo1 & "'", Conn
    If Not RSPelanggan1.EOF Then
        LblKodePelanggan = RSPelanggan1!kodePlg
        LblKodePelanggan.Enabled = False
        TxtKodeBarang.Enabled = True
    Else
        MsgBox "Nama Pelanggan tdak terdaftar"
        Combo1.SetFocus
    End If
End Sub

Private Sub HargaJual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtJumlah.SetFocus
End Sub

Private Sub KodeJasa_Change()

End Sub


Private Sub Option1_Click()
If Option1.Value = True Then
    Combo1.Enabled = False
    Combo1 = ""
    LblKodePelanggan = ""
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

Private Sub LblBiayaBarang_Change()
TotalBiaya = Format(Val(LblBiayaJasa) + Val(LblBiayaBarang), "###,###,###")
End Sub

Private Sub LblBiayaJasa_Change()
TotalBiaya = Format(Val(LblBiayaJasa) + Val(LblBiayaBarang), "###,###,###")
End Sub

Private Sub Timer1_Timer()
    LblJam = Time$
End Sub


'mencari nomor otomatis
Private Sub Auto()
Call BukaDB
RSService.Open "select * from Service Where Faktur In(Select Max(Faktur)From Service)Order By Faktur Desc", Conn
RSService.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSService
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

Function KosongkanJasa()

Call BukaDB
ADO1.ConnectionString = PathData
ADO1.RecordSource = "TempJasa"
ADO1.Refresh
If ADO1.Recordset.RecordCount <> 0 Then
    Do While Not ADO1.Recordset.EOF
        ADO1.Recordset.Delete
        ADO1.Recordset.MoveNext
    Loop
End If
LblBiayaJasa = 0
End Function

Private Sub CmdBatalJasa_Click()
Call KosongkanJasa
Call JasaLagi
End Sub


Function KosongkanBiaya()
'On Error Resume Next
Call BukaDB
ADO2.ConnectionString = PathData
ADO2.RecordSource = "Transaksi"
ADO2.Refresh
If ADO2.Recordset.RecordCount <> 0 Then
    Do While Not ADO2.Recordset.EOF
        ADO2.Recordset.Delete
        ADO2.Recordset.MoveNext
    Loop
End If
LblBiayaBarang = 0
End Function

Private Sub Bersihkan()
TxtNoPol = ""
LblNamaPelanggan = ""
Call KosongkanBiaya
Call KosongkanJasa
LblBiayaJasa = 0
Call Lagi
LblItem = 0
LblBiayaBarang = 0
TxtDibayar = 0
LblKembali = 0
TotalBiaya = 0
CBODiskon = "Diskon"
TxtDiskon = 0
TxtNoPol.SetFocus
End Sub

Private Sub TxtDibayar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtDibayar = "" Or Val(TxtDibayar) < (TotalBiaya) Then
            MsgBox "Jumlah Pembayaran Kurang"
            TxtDibayar.SetFocus
        Else
            TxtDibayar = Format(TxtDibayar, "###,###,###")
            If TxtDibayar = TotalBiaya Then
                LblKembali = TxtDibayar - TotalBiaya
            Else
                LblKembali = Format(TxtDibayar - TotalBiaya, "###,###,###")
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
If TxtNoPol = "" Or TotalBiaya = "" Or TotalBiaya = 0 Then
    MsgBox "Data belum lengkap"
    If TxtNoPol = "" Then
        TxtNoPol.SetFocus
    End If
End If

    
    Call BukaDB
    Dim JualTunai As String
    JualTunai = "Insert Into Service(Faktur,Tanggal,JmlItem,biayajasa,biayabarang,JmlTotal,Dibayar,Kembali,diskon,KodeKsr,nopol)" & _
    "values('" & TxtFaktur & "','" & Date & "','" & LblItem & "','" & LblBiayaJasa & "','" & LblBiayaBarang & "','" & TotalBiaya & "'," & _
    "'" & TxtDibayar & "','" & LblKembali & "','" & TxtDiskon & "','" & Menu.STBar.Panels(1).Text & "','" & TxtNoPol & "')"
    Conn.Execute (JualTunai)
    
    Dim Kas1 As String
    Kas1 = "insert into kas(tanggal,keterangan,pemasukan) values " & _
    "('" & Date & "','Jasa service dan Service barang','" & TotalBiaya & "')"
    Conn.Execute Kas1
    
    ADO1.Recordset.MoveFirst
    Do While Not ADO1.Recordset.EOF
        If ADO1.Recordset!kode <> vbNullString Then
            Dim DetailJasa As String
            DetailJasa = "Insert Into DetailJasa(Faktur,Kode_Jasa,nama_Jasa,harga) " & _
            "values ('" & TxtFaktur & "','" & ADO1.Recordset!kode & "','" & ADO1.Recordset!nama_jasa & "','" & ADO1.Recordset!Harga & "')"
            Conn.Execute (DetailJasa)
        End If
    ADO1.Recordset.MoveNext
    Loop
    
    ADO2.Recordset.MoveFirst
    Do While Not ADO2.Recordset.EOF
        If ADO2.Recordset!kode <> vbNullString Then
            Dim SQLTambahDetail As String
            SQLTambahDetail = "Insert Into DetailService(Faktur,Kodebrg,namabarang,hargaJual,JmlJual,subtotal) " & _
            "values ('" & TxtFaktur & "','" & ADO2.Recordset!kode & "','" & ADO2.Recordset!nama & "','" & ADO2.Recordset!Harga & "','" & ADO2.Recordset!jumlah & "','" & ADO2.Recordset!Total & "')"
            Conn.Execute (SQLTambahDetail)
        End If
    ADO2.Recordset.MoveNext
    Loop
        
    ADO2.Recordset.MoveFirst
    Do While Not ADO2.Recordset.EOF
            Call BukaDB
            RSBarang.Open "Select * from Barang where Kodebrg='" & ADO2.Recordset!kode & "'", Conn
            If Not RSBarang.EOF Then
                Dim KurangiStok As String
                KurangiStok = "update barang set jumlahbrg='" & RSBarang!jumlahbrg - ADO2.Recordset!jumlah & "' where kodebrg='" & ADO2.Recordset!kode & "'"
                Conn.Execute (KurangiStok)
            End If
    ADO2.Recordset.MoveNext
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
Call Bersihkan
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

'mencari total harga jasa
Function TotalHargaJasa()
    Dim RS1 As New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    RS1.Open "select sum(harga) as JumTotal from TempJasa", Conn
    TotalHargaJasa = RS1!JumTotal
End Function


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
            ADO2.Refresh
            DG2.Refresh
            LblBiayaBarang = TotalHarga
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
        'HargaJual.SetFocus
        TxtJumlah.SetFocus
        Exit Sub
    Else
        DaftarBarangService.Show
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
            ADO2.Refresh
            DG2.Refresh
            LblBiayaBarang = TotalHarga
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
            ADO2.Refresh
            DG2.Refresh
            LblBiayaBarang = TotalHarga
            LblItem = Format(TotalItem, "#,###,###")
            TxtKodeEdit.Visible = False
            TxtKodeEdit = ""
            Call Lagi
        End If
    End If
End If
End Sub

Private Sub TxtKodeJasa_KeyPress(KeyAscii As Integer)
TxtKodeJasa.MaxLength = 3
If KeyAscii = 27 Then Unload Me
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Call BukaDB
    RSJasa.Open "select * from jasa where kode_jasa='" & TxtKodeJasa & "'", Conn
    RSJasa.Requery
    If Not RSJasa.EOF Then
        NamaJasa = RSJasa!nama_jasa
        Harga = RSJasa!Harga
        '=====================================
        Dim Simpan As String
        Simpan = "insert into TempJasa(kode,Nama_jasa,harga) values " & _
        "('" & TxtKodeJasa & "','" & NamaJasa & "','" & Harga & "')"
        Conn.Execute Simpan
        Form_Activate
        ADO1.Refresh
        DG1.Refresh
        LblBiayaJasa = TotalHargaJasa
        Call JasaLagi
        Exit Sub
    Else
        DaftarJasa.Show
    End If
End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0

End Sub


Sub JasaLagi()
TxtKodeJasa = ""
NamaJasa = ""
Harga = ""
TxtKodeJasa.SetFocus
End Sub

Sub Lagi()
TxtKodeBarang = ""
NamaBarang = ""
HargaJual = ""
TxtJumlah = ""
Total = ""
TxtKodeBarang.SetFocus
End Sub


Function Cetak()
FakturService.Show
Call BukaDB
RSService.Open "select * from Service Where Faktur In(Select Max(Faktur)From Service)Order By Faktur Desc", Conn
Dim JmlHarga, JmlJual, JmlHasil As Double
Dim MGrs As String
FakturService.Font = "Courier New"
FakturService.Print
FakturService.Print
RSKasir.Open "select * From Kasir where KodeKsr= '" & RSService!kodeksr & "'", Conn
FakturService.Print Tab(5); "Faktur     :   "; RSService!faktur
FakturService.Print Tab(5); "Tanggal    :   "; Format(RSService!tanggal, "DD-MMMM-YYYY")
FakturService.Print Tab(5); "Kasir      :   "; RSKasir!NamaKsr
FakturService.Print Tab(5); "No Polisi  :   "; RSService!nopol
MGrs = String$(40, "-")
FakturService.Print Tab(5); MGrs
FakturService.Print Tab(5); "JASA / PELAYANAN"

RSDetailJasa.Open "select * from DetailJasa Where Faktur='" & RSService!faktur & "'", Conn
RSDetailJasa.MoveFirst
No = 0
Do While Not RSDetailJasa.EOF
    No = No + 1
    Harga = RSDetailJasa!Harga
    FakturService.Print Tab(5); No; Space(2); RSDetailJasa!nama_jasa;
    FakturService.Print Tab(30); RKanan(Harga, "###,###,###");
    RSDetailJasa.MoveNext
Loop
FakturService.Print Tab(5); MGrs
FakturService.Print Tab(5); "Total Biaya Jasa   :   ";
FakturService.Print Tab(30); RKanan(RSService!biayajasa, "###,###,###");

FakturService.Print Tab(5); MGrs
FakturService.Print Tab(5); "SPARE PART"

RSDetailService.Open "select * from DetailService Where Faktur='" & RSService!faktur & "'", Conn
RSDetailService.MoveFirst
No = 0
Do While Not RSDetailService.EOF
    No = No + 1
    Harga = RSDetailService!HargaJual
    jumlah = RSDetailService!JmlJual
    Hasil = Harga * jumlah

    FakturService.Print Tab(5); No; Space(2); RSDetailService!NamaBarang;
    FakturService.Print Tab(10); RKanan(jumlah, "##"); Space(1); "X";
    FakturService.Print Tab(15); RKanan(Harga, "###,###,###");
    FakturService.Print Tab(30); RKanan(Hasil, "###,###,###")
    RSDetailService.MoveNext
Loop
FakturService.Print Tab(5); MGrs
FakturService.Print Tab(5); "Total Biaya Barang : ";
FakturService.Print Tab(30); RKanan(RSService!biayabarang, "###,###,###");

FakturService.Print Tab(5); MGrs
If RSService!diskon = 0 Then
    FakturService.Print Tab(5); "Diskon    : ";
    FakturService.Print Tab(39); 0
Else
    FakturService.Print Tab(5); "Diskon    : ";
    FakturService.Print Tab(30); RKanan(RSService!diskon, "###,###,###");
End If

FakturService.Print Tab(5); "Total      :";
FakturService.Print Tab(30); RKanan(RSService!jmlTotal, "###,###,###");
FakturService.Print Tab(5); "Dibayar    :";

FakturService.Print Tab(30); RKanan(RSService!dibayar, "###,###,###");
FakturService.Print Tab(5); MGrs
FakturService.Print Tab(5); "Kembali    :";
If RSService!dibayar = RSService!jmlTotal Then
    FakturService.Print Tab(39); RSService!dibayar - RSService!jmlTotal
Else
    FakturService.Print Tab(30); RKanan(RSService!dibayar - RSService!jmlTotal, "###,###,###");
End If
FakturService.Print Tab(5); MGrs

FakturService.Print
Conn.Close
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function


Private Sub CBODiskon_Click()
TxtDiskon = LblBiayaBarang * Val(CBODiskon) / 100
LblBiayaBarang = LblBiayaBarang - Val(TxtDiskon)
End Sub

Private Sub TxtNoPol_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
    If TxtNoPol = "" Then
        LblNamaPelanggan = ""
        DataPendaftar.Show
        'MsgBox "Nopol polisi masih kosong"
        'TxtNoPol.SetFocus
    Else
        Call BukaDB
        RSPendaftaran.Open "select * from Pendaftaran where nopol='" & TxtNoPol & "'", Conn
        If Not RSPendaftaran.EOF Then
            LblNamaPelanggan = RSPendaftaran!nama
            TxtKodeJasa.SetFocus
        Else
            MsgBox "Nomor polisi tidak terdaftar"
            TxtNoPol.SetFocus
        End If
    End If
End If
End Sub
