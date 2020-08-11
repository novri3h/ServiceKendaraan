VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Aplikasi Service Kendaraan  Versi 1.0 [ Nadhif Studio ]"
   ClientHeight    =   7875
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   17145
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   7875
   ScaleWidth      =   17145
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   1395
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   19995
      _ExtentX        =   35269
      _ExtentY        =   2461
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   1058
      BackColor       =   16744576
      MouseIcon       =   "Menu.frx":214C0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Master"
      TabPicture(0)   =   "Menu.frx":214DC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image19"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Image4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Image3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Image2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Image1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   " Transaksi"
      TabPicture(1)   =   "Menu.frx":21DB6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image15"
      Tab(1).Control(1)=   "Image11"
      Tab(1).Control(2)=   "Image10"
      Tab(1).Control(3)=   "Image9"
      Tab(1).Control(4)=   "Image8"
      Tab(1).Control(5)=   "Image7"
      Tab(1).Control(6)=   "Image6"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   " Laporan"
      TabPicture(2)   =   "Menu.frx":22690
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image12"
      Tab(2).Control(1)=   "Image23"
      Tab(2).Control(2)=   "Image24"
      Tab(2).Control(3)=   "Image16"
      Tab(2).Control(4)=   "Image18"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   " Utility"
      TabPicture(3)   =   "Menu.frx":22F6A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Image26"
      Tab(3).Control(1)=   "Image27"
      Tab(3).Control(2)=   "Image21"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   " Keluar"
      TabPicture(4)   =   "Menu.frx":23844
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).ControlCount=   0
      Begin VB.Image Image18 
         Height          =   480
         Left            =   -73560
         Picture         =   "Menu.frx":23C96
         ToolTipText     =   "Rincian Penjualan"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image16 
         Height          =   480
         Left            =   -74280
         Picture         =   "Menu.frx":24560
         ToolTipText     =   "Rincian Pembelian"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image21 
         Height          =   345
         Left            =   -73680
         Picture         =   "Menu.frx":24E2A
         Top             =   840
         Width           =   360
      End
      Begin VB.Image Image19 
         Height          =   480
         Left            =   -70440
         Picture         =   "Menu.frx":25494
         ToolTipText     =   "Pendaftaran"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image17 
         Height          =   480
         Left            =   -69840
         Picture         =   "Menu.frx":25D5E
         ToolTipText     =   "Cari Barang"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image15 
         Height          =   480
         Left            =   -72360
         Picture         =   "Menu.frx":26628
         ToolTipText     =   "Service"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image14 
         Height          =   480
         Left            =   -71040
         Picture         =   "Menu.frx":26EF2
         ToolTipText     =   "Mekanik"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   -71760
         Picture         =   "Menu.frx":277BC
         ToolTipText     =   "Jasa"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image27 
         Height          =   480
         Left            =   -74280
         Picture         =   "Menu.frx":28086
         ToolTipText     =   "Backup Database"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image26 
         Height          =   480
         Left            =   -74880
         Picture         =   "Menu.frx":284C8
         ToolTipText     =   "Ganti Password User"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image24 
         Height          =   480
         Left            =   -72360
         Picture         =   "Menu.frx":2890A
         ToolTipText     =   "Laporan Stok Barang"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image23 
         Height          =   480
         Left            =   -72960
         Picture         =   "Menu.frx":28D4C
         ToolTipText     =   "Laporan Hutang Dan Piutang"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   -74880
         Picture         =   "Menu.frx":2918E
         ToolTipText     =   "Laporan Umum dan Transaksi"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   -71160
         Picture         =   "Menu.frx":29A58
         ToolTipText     =   "Penerimaan Piutang"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   -71760
         Picture         =   "Menu.frx":2A322
         ToolTipText     =   "Pembayaran Hutang"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   -73080
         Picture         =   "Menu.frx":2ABEC
         ToolTipText     =   "Retur Penjualan"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   -73680
         Picture         =   "Menu.frx":2B02E
         ToolTipText     =   "Penjualan"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   -74280
         Picture         =   "Menu.frx":2B470
         ToolTipText     =   "Retur Pembelian"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   -74880
         Picture         =   "Menu.frx":2B8B2
         ToolTipText     =   "Pembelian"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   -72480
         Picture         =   "Menu.frx":2C17C
         ToolTipText     =   "Pelanggan"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   -73080
         Picture         =   "Menu.frx":2CA46
         ToolTipText     =   "Pemasok"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -73680
         Picture         =   "Menu.frx":2D310
         ToolTipText     =   "Barang"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74280
         Picture         =   "Menu.frx":2DBDA
         ToolTipText     =   "Kasir"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -74880
         Picture         =   "Menu.frx":2E4A4
         ToolTipText     =   "Profil Perusahaan"
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   4200
   End
   Begin ComctlLib.StatusBar STBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   7380
      Width           =   17145
      _ExtentX        =   30242
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport CR 
      Left            =   600
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnfile 
      Caption         =   "Master"
      Begin VB.Menu mnprofil 
         Caption         =   "Profil Perusahaan"
      End
      Begin VB.Menu mnkasir 
         Caption         =   "Kasir"
      End
      Begin VB.Menu mnbarang 
         Caption         =   "Barang"
      End
      Begin VB.Menu mnpemasok 
         Caption         =   "Pemasok"
      End
      Begin VB.Menu mnpendaftaran 
         Caption         =   "Pendaftaran Service"
      End
      Begin VB.Menu mnpelangganbarang 
         Caption         =   "Pelanggan Barang"
      End
      Begin VB.Menu mnjasa 
         Caption         =   "Pekerjaan (Jasa)"
      End
      Begin VB.Menu mnmekanik 
         Caption         =   "Mekanik"
      End
      Begin VB.Menu mncaribarang 
         Caption         =   "Cari Data Barang"
      End
      Begin VB.Menu mnhapussemua 
         Caption         =   "Hapus isi semua tabel transaksi"
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnpembelian 
         Caption         =   "Pembelian"
      End
      Begin VB.Menu mnreturbeli 
         Caption         =   "Retur Pembelian"
      End
      Begin VB.Menu mnpenjualan 
         Caption         =   "Penjualan"
      End
      Begin VB.Menu MnReturJual 
         Caption         =   "Retur Penjualan"
      End
      Begin VB.Menu mnservice 
         Caption         =   "Service"
      End
      Begin VB.Menu mnbayarhutang 
         Caption         =   "Pembayaran Hutang"
      End
      Begin VB.Menu mnbayarpiutang 
         Caption         =   "Penerimaan Piutang"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlapmaster 
         Caption         =   "Data Master Dan Transaksi Umum"
      End
      Begin VB.Menu mnlaphutangpiutang 
         Caption         =   "Hutang Dan Piutang"
      End
      Begin VB.Menu mnrincibeli 
         Caption         =   "Rincian Pembelian"
      End
      Begin VB.Menu mnrincijual 
         Caption         =   "Rincian Penjualan"
      End
      Begin VB.Menu mnstokmin 
         Caption         =   "Stok Barang"
      End
   End
   Begin VB.Menu mnutility 
      Caption         =   "Utility"
      Begin VB.Menu mnganpass 
         Caption         =   "Ganti Password User"
      End
      Begin VB.Menu mnbackup 
         Caption         =   "Backup Database"
      End
      Begin VB.Menu mncetakfaktur 
         Caption         =   "Cetak Faktur"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
      Begin VB.Menu mnya 
         Caption         =   "Ya"
      End
      Begin VB.Menu mntidak 
         Caption         =   "Tidak"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
CariBarang.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Pesan = MsgBox("Tutup aplikasi ini..?", vbYesNo)
    If Pesan = vbYes Then End
End If
End Sub

Private Sub Form_Load()
Menu.STBar.Panels(4).Text = Date
If Menu.STBar.Panels(3) = "KASIR" Then
    Menu.SSTab1.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'End
End Sub

Private Sub Image1_Click()
Profil.Show
End Sub

Private Sub Image10_Click()
BayarHutang.Show
End Sub

Private Sub Image11_Click()
TerimaPiutang.Show
End Sub

Private Sub Image12_Click()
Laporan.Show
End Sub

Private Sub Image13_Click()
Jasa.Show
End Sub

Private Sub Image14_Click()
Mekanik.Show
End Sub

Private Sub Image15_Click()
Service.Show
End Sub

Private Sub Image16_Click()
RincianPembelian.Show
End Sub

Private Sub Image17_Click()
'LaporanReturPembelian.Show
CariBarang.Show
End Sub

Private Sub Image18_Click()
RincianPenjualan.Show
End Sub

Private Sub Image19_Click()
'LaporanReturPenjualan.Show
Pendaftaran.Show
End Sub

Private Sub Image2_Click()
Kasir.Show
End Sub

Private Sub Image20_Click()
'Laporan.Show
End Sub

Private Sub Image21_Click()
CetakFaktur.Show
End Sub

Private Sub Image22_Click()
LaporanHutang.Show
End Sub

Private Sub Image23_Click()
LaporanHutangPiutang.Show
End Sub

Private Sub Image24_Click()
StokMin.Show
End Sub

Private Sub Image25_Click()
'LaporanArusKas.Show
End Sub

Private Sub Image26_Click()
GantiPass.Show
End Sub

Private Sub Image27_Click()
BackupDatabase.Show
End Sub

Private Sub Image3_Click()
Barang.Show
End Sub

Private Sub Image4_Click()
Pemasok.Show
End Sub

Private Sub Image5_Click()
Pelanggan.Show
End Sub

Private Sub Image6_Click()
Pembelian.Show
End Sub

Private Sub Image7_Click()
ReturPembelian.Show
End Sub

Private Sub Image8_Click()
Penjualan.Show
End Sub

Private Sub Image9_Click()
ReturPenjualan.Show
End Sub

Private Sub mnbackup_Click()
BackupDatabase.Show
End Sub

Private Sub mnbarang_Click()
Barang.Show
End Sub

Private Sub mnbayarhutang_Click()
BayarHutang.Show
End Sub

Private Sub mnbayarpiutang_Click()
TerimaPiutang.Show
End Sub

Private Sub mndtbarang_Click()
    CR.DataFiles(0) = App.Path & "\dbretail.mdb"
    CR.ReportFileName = App.Path & "\Lap Barang.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mncaribarang_Click()
CariBarang.Show
End Sub

Private Sub mncetakfaktur_Click()
CetakFaktur.Show
End Sub

Private Sub mndtpembelian_Click()
'LaporanPembelian.Show
End Sub

Private Sub mndtrincian_Click()
Rincian.Show
End Sub

Private Sub mnganpass_Click()
GantiPass.Show
End Sub


Private Sub mnhapussemua_Click()
On Error Resume Next
Pesan = MsgBox("Yakin semua isi tabel transaksi akan dihapus", vbYesNo)
If Pesan = vbYes Then
    Call BukaDB
    Dim Hapus1 As String
    Hapus1 = "delete * from bayarhutang"
    Conn.Execute Hapus1
    
    Dim Hapus2 As String
    Hapus2 = "delete * from detailbeli"
    Conn.Execute Hapus2
    
    Dim Hapus3 As String
    Hapus3 = "delete * from detailjual"
    Conn.Execute Hapus3
    
    Dim Hapus4 As String
    Hapus4 = "delete * from detailreturbeli"
    Conn.Execute Hapus4
    
    Dim Hapus5 As String
    Hapus5 = "delete * from detailreturjual"
    Conn.Execute Hapus5
    
    Dim Hapus6 As String
    Hapus6 = "delete * from kas"
    Conn.Execute Hapus6
    
    Dim Hapus7 As String
    Hapus7 = "delete * from pembelian"
    Conn.Execute Hapus7
    
    Dim Hapus8 As String
    Hapus8 = "delete * from penjualan"
    Conn.Execute Hapus8
    
    Dim Hapus9 As String
    Hapus9 = "delete * from returbeli"
    Conn.Execute Hapus9
    
    Dim Hapus10 As String
    Hapus10 = "delete * from returjual"
    Conn.Execute Hapus10
    
    Dim Hapus11 As String
    Hapus11 = "delete * from terimapiutang"
    Conn.Execute Hapus11
    
    Dim Hapus12 As String
    Hapus12 = "delete * from service"
    Conn.Execute Hapus12
    
    Dim Hapus13 As String
    Hapus13 = "delete * from detailservice"
    Conn.Execute Hapus13
    
    Dim Hapus14 As String
    Hapus14 = "delete * from detailjasa"
    Conn.Execute Hapus14
    
    Dim Hapus15 As String
    Hapus15 = "delete * from pendaftaran"
    Conn.Execute Hapus15
    
    'Dim nolkan As String
    'nolkan = "update barang set jumlahbrg=0"
    'Conn.Execute nolkan
    
    MsgBox "Penghapusan telah dilakukan dan tidak dapat dibatalkan"
End If
End Sub

Private Sub mnjasa_Click()
Jasa.Show
End Sub

Private Sub mnkasir_Click()
Kasir.Show
End Sub

Private Sub mnlapbayarhutang_Click()
LaporanBayarHutang.Show
End Sub

Private Sub mnlapbayarpiutang_Click()
LaporanTerimaPiutang.Show
End Sub

Private Sub mnlaphutang_Click()
LaporanHutang.Show
End Sub

Private Sub mnlappelanggan_Click()
CR.DataFiles(0) = App.Path & "\dbretail.mdb"
    CR.ReportFileName = App.Path & "\Lap Pelanggan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnlappemakai_Click()
'LaporanUmum.Show

CR.DataFiles(0) = App.Path & "\dbretail.mdb"
    CR.ReportFileName = App.Path & "\Lap Pemakai.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnlappemasok_Click()
    CR.DataFiles(0) = App.Path & "\dbretail.mdb"
    CR.ReportFileName = App.Path & "\Lap Pemasok.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnlaphutangpiutang_Click()
LaporanHutangPiutang.Show
End Sub

Private Sub mnlapmaster_Click()
Laporan.Show
End Sub

Private Sub mnlappenjualan_Click()
'LaporanPenjualan.Show
End Sub

Private Sub mnlappiutang_Click()
LaporanPiutang.Show
End Sub

Private Sub mnlapreturbeli_Click()
LaporanReturPembelian.Show
End Sub

Private Sub mnlapreturjual_Click()
LaporanReturPenjualan.Show
End Sub

Private Sub mnmekanik_Click()
Mekanik.Show
End Sub

Private Sub mnpelanggan_Click()
Pelanggan1.Show
End Sub

Private Sub mnpelangganbarang_Click()
Pelanggan.Show
End Sub

Private Sub mnpemasok_Click()
Pemasok.Show
End Sub

Private Sub mnpembelian_Click()
'Pembelian.Show
Pembelian.Show
End Sub

Private Sub mnpendaftaran_Click()
Pendaftaran.Show
End Sub

Private Sub mnpenjualan_Click()
Penjualan.Show
End Sub

Private Sub mnprofil_Click()
Profil.Show
End Sub

Private Sub mnreturbeli_Click()
ReturPembelian.Show
End Sub

Private Sub mnreturjual_Click()
ReturPenjualan.Show
End Sub

Private Sub mnrincian_Click()
'Rincian.Show
End Sub

Private Sub mnuji_Click()
UjiSQL.Show
End Sub

Private Sub mnrincibeli_Click()
RincianPembelian.Show
End Sub

Private Sub mnrincijual_Click()
RincianPenjualan.Show
End Sub

Private Sub mnsql_Click()
UjiSQL.Show
End Sub

Private Sub mnservice_Click()
Service.Show
End Sub

Private Sub mnstokmin_Click()
StokMin.Show
End Sub

Private Sub mnujisql_Click()
UjiSQL.Show
End Sub

Private Sub mnya_Click()
End
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next

If SSTab1.Tab = 4 Then
    Conn.Close
    Pesan = MsgBox("Tutup aplikasi ini..?", vbYesNo)
    If Pesan = vbYes Then End
End If
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Pesan = MsgBox("Tutup aplikasi ini..?", vbYesNo)
    If Pesan = vbYes Then End
End If
End Sub

Private Sub Timer1_Timer()
Menu.STBar.Panels(5).Text = Time$
End Sub
