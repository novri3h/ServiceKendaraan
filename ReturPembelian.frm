VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ReturPembelian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retur Pembelian"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
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
   ScaleHeight     =   6930
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4470
      TabIndex        =   41
      Top             =   540
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
      Left            =   5430
      TabIndex        =   40
      Top             =   570
      Width           =   735
   End
   Begin VB.TextBox TxtFaktur 
      Height          =   350
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   240
      TabIndex        =   9
      Top             =   5760
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
      TabIndex        =   8
      Top             =   5760
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3960
      Top             =   6240
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
      TabIndex        =   7
      Top             =   960
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
      Top             =   960
      Width           =   885
   End
   Begin VB.TextBox TxtKodeBatal 
      Height          =   350
      Left            =   1440
      TabIndex        =   4
      Text            =   "Kode Barang ?"
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit Jumlah"
      Height          =   350
      Left            =   2640
      TabIndex        =   3
      Top             =   5760
      Width           =   1200
   End
   Begin VB.CommandButton CmdHapusSemua 
      Caption         =   "&Batalkan Semua Item Retur"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   3360
      Width           =   2100
   End
   Begin VB.TextBox TxtKodeEdit 
      Height          =   350
      Left            =   2640
      TabIndex        =   1
      Text            =   "Kode Barang ?"
      Top             =   6240
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "ReturPembelian.frx":0000
      Height          =   1800
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3175
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
            ColumnWidth     =   1454,74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   345
      Left            =   2280
      Top             =   3360
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
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
   Begin MSDataGridLib.DataGrid DGRetur 
      Bindings        =   "ReturPembelian.frx":0012
      Height          =   1800
      Left            =   120
      TabIndex        =   27
      Top             =   1440
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3175
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
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
      Caption         =   "Barang yang pernah dibeli untuk dikembalikan (Retur)"
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
            ColumnWidth     =   1454,74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADORetur 
      Height          =   345
      Left            =   120
      Top             =   3360
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "ADORetur"
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
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jenis Pembelian"
      Height          =   345
      Left            =   2280
      TabIndex        =   39
      Top             =   480
      Width           =   2050
   End
   Begin VB.Label LblNomorRetur 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   120
      TabIndex        =   38
      Top             =   480
      Width           =   2000
   End
   Begin VB.Label LblKodePemasok 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   8760
      TabIndex        =   37
      Top             =   120
      Width           =   885
   End
   Begin VB.Label LblNamaPemasok 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5400
      TabIndex        =   36
      Top             =   120
      Width           =   3330
   End
   Begin VB.Label LblTempo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   8520
      TabIndex        =   35
      Top             =   5760
      Width           =   1245
   End
   Begin VB.Label LblUangMuka 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   8520
      TabIndex        =   34
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label LblDibayar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6000
      TabIndex        =   33
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Retur"
      Height          =   345
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   2000
   End
   Begin VB.Label LblItemRetur 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6720
      TabIndex        =   31
      Top             =   3360
      Width           =   600
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item"
      Height          =   345
      Left            =   6120
      TabIndex        =   30
      Top             =   3360
      Width           =   600
   End
   Begin VB.Label LblTotalRetur 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   8160
      TabIndex        =   29
      Top             =   3360
      Width           =   1245
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   345
      Left            =   7320
      TabIndex        =   28
      Top             =   3360
      Width           =   795
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tempo"
      Height          =   345
      Left            =   7440
      TabIndex        =   26
      Top             =   5760
      Width           =   1005
   End
   Begin VB.Label LblSisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   8520
      TabIndex        =   25
      Top             =   6480
      Width           =   1245
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa"
      Height          =   345
      Left            =   7440
      TabIndex        =   24
      Top             =   6480
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
      TabIndex        =   23
      Top             =   6120
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Faktur"
      Height          =   345
      Left            =   2280
      TabIndex        =   22
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   345
      Left            =   5160
      TabIndex        =   21
      Top             =   5760
      Width           =   795
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6000
      TabIndex        =   20
      Top             =   5760
      Width           =   1245
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   5160
      TabIndex        =   19
      Top             =   6120
      Width           =   795
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   345
      Left            =   5160
      TabIndex        =   18
      Top             =   6480
      Width           =   795
   End
   Begin VB.Label LblKembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6000
      TabIndex        =   17
      Top             =   6480
      Width           =   1245
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item"
      Height          =   345
      Left            =   3960
      TabIndex        =   16
      Top             =   5760
      Width           =   600
   End
   Begin VB.Label LblItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4560
      TabIndex        =   15
      Top             =   5760
      Width           =   600
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pemasok"
      Height          =   345
      Left            =   4440
      TabIndex        =   14
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   960
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
      TabIndex        =   12
      Top             =   960
      Width           =   3435
   End
   Begin VB.Label HargaBeli 
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
      TabIndex        =   11
      Top             =   960
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
      TabIndex        =   10
      Top             =   960
      Width           =   1500
   End
End
Attribute VB_Name = "ReturPembelian"
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

ADORetur.ConnectionString = PathData '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBeli.mdb"
ADORetur.RecordSource = "Transaksi"
Set DGRetur.DataSource = ADORetur
DGRetur.Refresh
Call Auto
End Sub

Private Sub Auto()
Call BukaDB
RSReturBeli.Open "select * from ReturBeli Where NomorRetur In(Select Max(NomorRetur)From ReturBeli)Order By NomorRetur Desc", Conn
RSReturBeli.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSReturBeli
        If .EOF Then
            Urutan = Format(Date, "yymmdd") + "0001"
            LblNomorRetur = Urutan
        Else
            If Left(!nomorretur, 6) <> Format(Date, "yymmdd") Then
                Urutan = Format(Date, "yymmdd") + "0001"
            Else
                Hitung = (!nomorretur) + 1
                Urutan = Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
            End If
        End If
        LblNomorRetur = Urutan
    End With
End Sub


Private Sub Form_Load()
Call Tabel_Kosong
Call BukaDB
Option1.Value = True
TxtKodeBatal.Visible = False
TxtKodeEdit.Visible = False
Option1.Enabled = False
Option2.Enabled = False
TxtKodeBarang.Enabled = False
End Sub


Function Tabel_Kosong()
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
LblNamaPemasok = ""
LblKodePemasok = ""
TxtKodeBarang = ""
Call Tabel_Kosong
LblItem = ""
LblTotal = ""
LblDibayar = ""
LblKembali = ""
LblTempo = ""
LblUangMuka = ""
LblSisa = ""
LblItemRetur = ""
LblTotalRetur = ""
TxtFaktur = ""
TxtFaktur.SetFocus
End Sub

Private Sub CmdSimpan_Click()
If TxtFaktur = "" Or LblItemRetur = "" Or LblTotalRetur = "" Then
    MsgBox "Data belum lengkap"
    Exit Sub
Else
    Call BukaDB
    Dim SimpanRetur As String
    SimpanRetur = "insert into ReturBeli(nomorretur,faktur,kodeksr,kodepms,tanggalretur,itemretur,totalretur) values " & _
    "('" & LblNomorRetur & "','" & TxtFaktur & "','" & Menu.STBar.Panels(1) & "','" & LblKodePemasok & "','" & Menu.STBar.Panels(4) & "','" & LblItemRetur & "','" & LblTotalRetur & "')"
    Conn.Execute SimpanRetur

    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF
        Dim SimpanDetailRetur As String
        SimpanDetailRetur = "insert into DetailReturBeli(nomorretur,faktur,kodebrg,namabarang,hargabeli,jmlretur,totalretur) values " & _
        "('" & LblNomorRetur & "','" & TxtFaktur & "','" & ADO.Recordset!kode & "','" & ADO.Recordset!nama & "','" & ADO.Recordset!Harga & "','" & ADO.Recordset!jumlah & "','" & ADO.Recordset!Total & "')"
        Conn.Execute (SimpanDetailRetur)
    ADO.Recordset.MoveNext
    Loop
    
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF
        Call BukaDB
        RSBarang.Open "Select * from Barang where Kodebrg='" & ADO.Recordset!kode & "'", Conn
        If Not RSBarang.EOF Then
            Dim TambahBarang As String
            TambahBarang = "update barang set jumlahbrg='" & RSBarang!jumlahbrg - ADO.Recordset!jumlah & "' where kodebrg='" & ADO.Recordset!kode & "'"
            Conn.Execute (TambahBarang)
        End If
    ADO.Recordset.MoveNext
    Loop
    
    ADO.Recordset.MoveFirst
    Do While Not ADO.Recordset.EOF
        Call BukaDB
        RSDetailBeli.Open "Select * from detailbeli where Kodebrg='" & ADO.Recordset!kode & "' and faktur='" & TxtFaktur & "'", Conn
        If Not RSDetailBeli.EOF Then
            Dim tambahdetail As String
            tambahdetail = "update detailbeli set jmlbeli='" & RSDetailBeli!Jmlbeli - ADO.Recordset!jumlah & "' where kodebrg='" & ADO.Recordset!kode & "' and faktur='" & TxtFaktur & "'"
            Conn.Execute (tambahdetail)
        End If
    ADO.Recordset.MoveNext
    Loop
    
    Dim Kas As String
    Kas = "insert into kas(tanggal,keterangan,pemasukan) values " & _
    "('" & Date & "','" & ReturPembelian.Caption + Space(1) + LblNamaPemasok & "','" & LblTotalRetur & "')"
    Conn.Execute Kas


    Bersihkan
    Form_Activate
End If
ADORetur.RecordSource = "select kodebrg as kode,namabarang as nama,hargabeli as harga,jmlbeli as jumlah,subtotal as total from detailbeli where faktur='xx'"
ADORetur.Refresh
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
    TxtKodeBarang = ""
    NamaBarang = ""
    HargaBeli = ""
    TxtJumlah = ""
    TxtKodeBatal.Visible = False
    TxtKodeEdit.Visible = False
    Call Tabel_Kosong
    LblItemRetur = ""
    LblTotalRetur = ""
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

Private Sub TxtFaktur_Change()
If TxtFaktur = "" Then
    ADORetur.RecordSource = "select kodebrg as kode,namabarang as nama,hargabeli as harga,jmlbeli as jumlah,subtotal as total from detailbeli where faktur='" & TxtFaktur & "'"
    ADORetur.Refresh
    LblNamaPemasok = ""
    LblKodePemasok = ""
    LblItem = ""
    LblTotal = ""
    LblDibayar = ""
    LblKembali = ""
    LblTempo = ""
    LblUangMuka = ""
    LblSisa = ""
    TxtKodeBarang.Enabled = False
'Else
'    TxtFaktur_KeyPress (13)
End If
End Sub

Private Sub TxtFaktur_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
    If TxtFaktur = "" Then
        DaftarFaktur.Show
        Exit Sub
    Else
        ADORetur.RecordSource = "select kodebrg as kode,namabarang as nama,hargabeli as harga,jmlbeli as jumlah,subtotal as total from detailbeli where faktur='" & TxtFaktur & "'"
        ADORetur.Refresh
        TxtKodeBarang.Enabled = True
        Call BukaDB
        RSPembelian.Open "select * from pembelian where faktur='" & TxtFaktur & "'", Conn
        If RSPembelian.EOF Then
            MsgBox "Nomor faktur tidak terdaftar"
            TxtFaktur.SetFocus
            Exit Sub
        Else
            ADORetur.RecordSource = "select kodebrg as kode,namabarang as nama,hargabeli as harga,jmlbeli as jumlah,subtotal as total from detailbeli where faktur='" & TxtFaktur & "'"
            ADORetur.Refresh
            LblKodePemasok = RSPembelian!kodepms
            If RSPembelian!jenis = "Tunai" Then
                Option1.Value = True
                Option1.Enabled = True
                Option2.Enabled = False
                LblDibayar = RSPembelian!dibayar
                LblKembali = RSPembelian!kembali
                LblTempo = ""
                LblUangMuka = ""
                LblSisa = ""
            Else
                Option2.Value = True
                Option2.Enabled = True
                Option1.Enabled = False
                LblTempo = RSPembelian!tempo
                LblUangMuka = RSPembelian!DP
                LblSisa = RSPembelian!sisa
                LblDibayar = ""
                LblKembali = ""
            End If
            LblTotal = RSPembelian!jmlTotal
            LblItem = RSPembelian!jmlitem
            RSPemasok.Open "select * from pemasok where kodepms='" & RSPembelian!kodepms & "'", Conn
            LblNamaPemasok = RSPemasok!namapms
        End If
    End If
End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub TxtJumlah_KeyPress(KeyAscii As Integer)
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
        LblTotalRetur = Format(TotalHarga, "###,###,###")
        LblItemRetur = Format(TotalItem, "#,###,###")
        Call Lagi
    End If
End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub TxtKodeBarang_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Call BukaDB
    RSDetailBeli.Open "select * from detailbeli where kodebrg='" & TxtKodeBarang & "' and faktur='" & TxtFaktur & "'", Conn
    RSDetailBeli.Requery
    If Not RSDetailBeli.EOF Then
        NamaBarang = RSDetailBeli!NamaBarang
        HargaBeli = RSDetailBeli!HargaBeli
        TxtJumlah = RSDetailBeli!Jmlbeli
        TxtJumlah.SetFocus
        Exit Sub
    Else
        MsgBox "Kode tidak ditemukan, lihat kode di dalam list retur, atau mungkin faktur belum diisi"
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


Sub Lagi()
TxtKodeBarang = ""
NamaBarang = ""
HargaBeli = ""
TxtJumlah = ""
Total = ""
TxtKodeBarang.SetFocus
End Sub

