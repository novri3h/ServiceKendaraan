VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form LaporanHutangPiutang 
   Caption         =   "Laporan Hutang Piutang"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Laporan Hutang"
      TabPicture(0)   =   "LaporanHutangPiutang.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CR"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "List1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Combo1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Laporan Piutang"
      TabPicture(1)   =   "LaporanHutangPiutang.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "CrystalReport1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Combo2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "List2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.ListBox List2 
         Height          =   2310
         Left            =   2520
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   2520
         TabIndex        =   3
         Text            =   "Pilih Data...!"
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   -74880
         TabIndex        =   2
         Text            =   "Pilih Data...!"
         Top             =   600
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Height          =   2310
         Left            =   -74880
         TabIndex        =   1
         Top             =   1080
         Width           =   2295
      End
      Begin Crystal.CrystalReport CR 
         Left            =   -73920
         Top             =   2040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   1080
         Top             =   2040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "LaporanHutangPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Combo1.AddItem "Hutang Jatuh Tempo"
Combo1.AddItem "Status Hutang"
Combo1.AddItem "Hutang Bulanan"
Combo1.AddItem "Pembayaran Hutang Bulanan"

Combo2.AddItem "Piutang Jatuh Tempo"
Combo2.AddItem "Status Piutang"
Combo2.AddItem "Piutang Bulanan"
Combo2.AddItem "Penerimaan Piutang Bulanan"
End Sub


Private Sub Combo1_Click()
If Combo1 = "Hutang Jatuh Tempo" Then
    List1.Clear
    Call BukaDB
    RSPembelian.Open "select distinct jatuhtempo from pembelian where jenis='kredit' and sisa<>0", Conn
    If Not RSPembelian.EOF Then
        Do While Not RSPembelian.EOF
            List1.AddItem Format(RSPembelian!jatuhtempo, "DD-MMM-YYYY")
            RSPembelian.MoveNext
        Loop
    Else
        MsgBox "Belum ada data hutang jatuh tempo"
    End If
ElseIf Combo1 = "Status Hutang" Then
    List1.Clear
    List1.AddItem "Lunas"
    List1.AddItem "Belum Lunas"
ElseIf Combo1 = "Hutang Bulanan" Then
    List1.Clear
    Call BukaDB
    RSPembelian.Open "SELECT DISTINCT MONTH(JATUHTEMPO) as bulan,YEAR(JATUHTEMPO) as tahun FROM PEMBELIAN WHERE JATUHTEMPO<>0", Conn
    If Not RSPembelian.EOF Then
        Do While Not RSPembelian.EOF
            List1.AddItem RSPembelian!Bulan & Space(1) & MonthName(RSPembelian!Bulan) & vbTab & RSPembelian!Tahun
            RSPembelian.MoveNext
        Loop
    Else
        MsgBox "Belum ada data"
    End If
ElseIf Combo1 = "Pembayaran Hutang Bulanan" Then
    List1.Clear
    Call BukaDB
    RSHutang.Open "SELECT DISTINCT MONTH(tanggalbayar) as bulan,YEAR(tanggalbayar) as tahun FROM bayarhutang", Conn
    If Not RSHutang.EOF Then
        Do While Not RSHutang.EOF
            List1.AddItem RSHutang!Bulan & Space(1) & MonthName(RSHutang!Bulan) & vbTab & RSHutang!Tahun
            RSHutang.MoveNext
        Loop
    Else
        MsgBox "Belum ada data"
    End If
End If
    
End Sub


Private Sub List1_Click()
If Combo1 = "Hutang Jatuh Tempo" Then
    CR.SelectionFormula = "Totext({Pembelian.JatuhTempo})='" & CDate(List1) & "' AND {PEMBELIAN.SISA}<>0"
    CR.ReportFileName = App.Path & "\Lap hutang jt hARIAN.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
ElseIf List1 = "Lunas" Then
    Call BukaDB
    RSPembelian.Open "select * from pembelian where jenis='kredit' and sisa=0", Conn
    If RSPembelian.EOF Then
        MsgBox "Belum ada data"
        Exit Sub
    Else
        CR.SelectionFormula = "{Pembelian.sisa}=0" ' and {bayarhutang.keterangan}='LUNAS'"
        CR.ReportFileName = App.Path & "\Lap hutang status.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
        CR.Reset
    End If
ElseIf List1 = "Belum Lunas" Then
    Call BukaDB
    RSPembelian.Open "select * from pembelian where jenis='kredit' and sisa<>0", Conn
    If RSPembelian.EOF Then
        MsgBox "Belum ada data"
        Exit Sub
    Else
        CR.SelectionFormula = "{Pembelian.sisa}<>0"
        CR.ReportFileName = App.Path & "\Lap hutang status.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
        CR.Reset
    End If
ElseIf Combo1 = "Hutang Bulanan" Then
    CR.SelectionFormula = "Month({Pembelian.JatuhTempo})=" & Val(Trim(Left(List1, 2))) & " and Year({pembelian.JatuhTempo})=" & Val(Right(List1, 4)) & " and {pembelian.sisa}<>0"
    CR.ReportFileName = App.Path & "\Lap hutang Bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
ElseIf Combo1 = "Pembayaran Hutang Bulanan" Then
    CR.SelectionFormula = "Month({bayarhutang.tanggalbayar})=" & Val(Trim(Left(List1, 2))) & " and Year({bayarhutang.tanggalbayar})=" & Val(Right(List1, 4)) & " and {pembelian.jenis}='Kredit'"
    CR.ReportFileName = App.Path & "\Lap hutang Bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End If

End Sub

'=========================================

Private Sub Combo2_Click()
If Combo2 = "Piutang Jatuh Tempo" Then
    List2.Clear
    Call BukaDB
    RSPenjualan.Open "select distinct jatuhtempo from Penjualan where jenis='kredit' and sisa<>0", Conn
    If Not RSPenjualan.EOF Then
        Do While Not RSPenjualan.EOF
            List2.AddItem Format(RSPenjualan!jatuhtempo, "DD-MMM-YYYY")
            RSPenjualan.MoveNext
        Loop
    Else
        MsgBox "Belum ada data"
    End If
ElseIf Combo2 = "Status Piutang" Then
    List2.Clear
    List2.AddItem "Lunas"
    List2.AddItem "Belum Lunas"
ElseIf Combo2 = "Piutang Bulanan" Then
    List2.Clear
    Call BukaDB
    RSPenjualan.Open "SELECT DISTINCT MONTH(JATUHTEMPO) as bulan,YEAR(JATUHTEMPO) as tahun FROM Penjualan WHERE SISA<>0", Conn
    If Not RSPenjualan.EOF Then
        Do While Not RSPenjualan.EOF
            List2.AddItem RSPenjualan!Bulan & Space(1) & MonthName(RSPenjualan!Bulan) & vbTab & RSPenjualan!Tahun
            RSPenjualan.MoveNext
        Loop
    Else
        MsgBox "Belum ada data"
    End If
ElseIf Combo2 = "Penerimaan Piutang Bulanan" Then
    List2.Clear
    Call BukaDB
    RSPiutang.Open "SELECT DISTINCT MONTH(tanggalterima) as bulan,YEAR(tanggalterima) as tahun FROM terimapiutang", Conn ' WHERE SISA<>0", Conn
    If Not RSPiutang.EOF Then
        Do While Not RSPiutang.EOF
            List2.AddItem RSPiutang!Bulan & Space(1) & MonthName(RSPiutang!Bulan) & vbTab & RSPiutang!Tahun
            RSPiutang.MoveNext
        Loop
    Else
        MsgBox "Belum ada data"
    End If

End If
    
End Sub


Private Sub List2_Click()
If Combo2 = "Piutang Jatuh Tempo" Then
    CR.SelectionFormula = "Totext({Penjualan.JatuhTempo})='" & CDate(List2) & "' AND {Penjualan.SISA}<>0"
    CR.ReportFileName = App.Path & "\Lap Piutang jatuh tempo.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
ElseIf List2 = "Lunas" Then
    CR.SelectionFormula = "{Penjualan.sisa}=0"
    CR.ReportFileName = App.Path & "\Lap Piutang status.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
ElseIf List2 = "Belum Lunas" Then
    CR.SelectionFormula = "{Penjualan.sisa}<>0"
    CR.ReportFileName = App.Path & "\Lap Piutang status.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset

ElseIf Combo2 = "Piutang Bulanan" Then
    CR.SelectionFormula = "Month({Penjualan.JatuhTempo})=" & Val(Trim(Left(List2, 2))) & " and Year({Penjualan.JatuhTempo})=" & Val(Right(List2, 4)) & " and {Penjualan.sisa}<>0"
    CR.ReportFileName = App.Path & "\Lap Piutang Bulanan1.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
ElseIf Combo2 = "Penerimaan Piutang Bulanan" Then
    CR.SelectionFormula = "Month({terimapiutang.tanggalterima})=" & Val(Trim(Left(List2, 2))) & " and Year({terimapiutang.tanggalterima})=" & Val(Right(List2, 4)) '& " and {terimapiutang.sisa}<>0"
    CR.ReportFileName = App.Path & "\Lap Piutang Bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
    
End If

End Sub


