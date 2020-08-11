VERSION 5.00
Begin VB.Form CetakFaktur 
   Caption         =   "Cetak Faktur"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3525
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
   ScaleHeight     =   2565
   ScaleWidth      =   3525
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   1635
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1635
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Penjualan Kredit"
      Height          =   225
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Penjualan Tunai"
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "CetakFaktur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call BukaDB
RSPenjualan.Open "select * from penjualan where jenis='Tunai'", Conn
List1.Clear
Do While Not RSPenjualan.EOF
    List1.AddItem RSPenjualan!faktur
    RSPenjualan.MoveNext
Loop
Conn.Close

Call BukaDB
RSPenjualan.Open "select * from penjualan where jenis='Kredit'", Conn
List2.Clear
Do While Not RSPenjualan.EOF
    List2.AddItem RSPenjualan!faktur
    RSPenjualan.MoveNext
Loop
Conn.Close

End Sub


Function Cetak()
Dim JmlHarga, JmlJual, JmlHasil As Double
Dim MGrs As String
Layar.Font = "Courier New"
Layar.Print
Layar.Print
RSKasir.Open "select * From Kasir where KodeKsr= '" & RSPenjualan!kodeksr & "'", Conn
If RSPenjualan!jenis = "Tunai" Then
    Layar.Print Tab(5); "Faktur     :   "; RSPenjualan!faktur
    Layar.Print Tab(5); "Tanggal    :   "; Format(RSPenjualan!tanggal, "DD-MMMM-YYYY")
    Layar.Print Tab(5); "Kasir      :   "; RSKasir!NamaKsr
ElseIf RSPenjualan!jenis = "Kredit" Then
    RSPelanggan.Open "select * From Pelanggan where KodePlg= '" & RSPenjualan!kodePlg & "'", Conn
    Layar.Print Tab(5); "Faktur     :   "; RSPenjualan!faktur
    Layar.Print Tab(5); "Tanggal    :   "; Format(RSPenjualan!tanggal, "DD-MMMM-YYYY")
    Layar.Print Tab(5); "Kasir      :   "; RSKasir!NamaKsr
    Layar.Print Tab(5); "Jenis      :   "; RSPenjualan!jenis
    Layar.Print Tab(5); "Pelanggan  :   "; RSPelanggan!NamaPlg
    Layar.Print Tab(5); "Telepon    :   "; RSPelanggan!teleponPlg
End If
MGrs = String$(33, "-")
Layar.Print Tab(5); MGrs

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
    'Layar berulang-ulang kode,nama,harga,jumlah dan total
    Layar.Print Tab(5); No; Space(2); RSBarang!namabrg
    Layar.Print Tab(10); RKanan(jumlah, "##"); Space(1); "X";
    Layar.Print Tab(15); Format(Harga, "###,###,###");
    Layar.Print Tab(25); RKanan(Hasil, "###,###,###")
    RSDetailJual.MoveNext
Loop

'Layar total harga
If RSPenjualan!jenis = "Tunai" Then
    Layar.Print Tab(5); MGrs
    Layar.Print Tab(5); "Total      :";
    Layar.Print Tab(25); RKanan(RSPenjualan!jmlTotal, "###,###,###");
    Layar.Print Tab(5); "Dibayar    :";
    'Layar dibayar
    Layar.Print Tab(25); RKanan(RSPenjualan!dibayar, "###,###,###");
    Layar.Print Tab(5); MGrs
    Layar.Print Tab(5); "Kembali    :";
    'Layar kembalian
    If RSPenjualan!dibayar = RSPenjualan!jmlTotal Then
        Layar.Print Tab(34); RSPenjualan!dibayar - RSPenjualan!jmlTotal
    Else
        Layar.Print Tab(25); RKanan(RSPenjualan!dibayar - RSPenjualan!jmlTotal, "###,###,###");
    End If
    Layar.Print Tab(5); MGrs
ElseIf RSPenjualan!jenis = "Kredit" Then
    Layar.Print Tab(5); MGrs
    Layar.Print Tab(5); "Total      :";
    Layar.Print Tab(25); RKanan(RSPenjualan!jmlTotal, "###,###,###");
    Layar.Print Tab(5); "Tempo      :";
    Layar.Print Tab(25); RKanan(RSPenjualan!tempo, "###,###,###");
    Layar.Print Tab(5); "Jatuh Tempo:";
    Layar.Print Tab(25); RKanan(RSPenjualan!jatuhtempo, "dd-mmm-yyyy");
    Layar.Print Tab(5); "Uang Muka  :";
    Layar.Print Tab(25); RKanan(RSPenjualan!DP, "###,###,###");
    Layar.Print Tab(5); "Sisa       :";
    Layar.Print Tab(25); RKanan(RSPenjualan!sisa, "###,###,###");
    If RSPenjualan!sisa = 0 Then
        Layar.Print Tab(5); "Keterangan : Lunas";
        'Layar.Print Tab(25); "Lunas";  ''RKanan(RSPenjualan!jatuhtempo, "; dd - mmm - yyyy; ");"
    End If
    
    Layar.Print Tab(5); MGrs
End If
Layar.Print
Layar.Print
Layar.Print
'Layar.EndDoc
Conn.Close
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function



Private Sub List1_Click()
Layar.Show
Layar.Caption = CetakFaktur.List1
Call BukaDB
RSPenjualan.Open "select * from Penjualan Where Faktur ='" & List1 & "'", Conn
Call Cetak
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub List2_Click()
Layar.Show
Layar.Caption = CetakFaktur.List2
Call BukaDB
RSPenjualan.Open "select * from Penjualan Where Faktur ='" & List2 & "'", Conn
Call Cetak
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
