VERSION 5.00
Begin VB.Form Tampilkan 
   BackColor       =   &H80000009&
   Caption         =   "ESC = Tutup ** Enter = Cetak"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
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
   ScaleHeight     =   5730
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Tampilkan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
ElseIf KeyAscii = 13 Then
    Pesan = MsgBox("Printer sudah siap", vbYesNo)
    If Pesan = vbYes Then
        Call Cetak
    Else
        Unload Me
    End If
End If
End Sub

Function Cetak()
Call BukaDB
'cari faktur terakhir
RSPenjualan.Open "select * from Penjualan Where Faktur In(Select Max(Faktur)From Penjualan)Order By Faktur Desc", Conn
Dim JmlHarga, JmlJual, JmlHasil As Double
Dim MGrs As String
Printer.Font = "Courier New"
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print
Printer.Print
Printer.Print
RSKasir.Open "select * From Kasir where KodeKsr= '" & RSPenjualan!kodeksr & "'", Conn
If RSPenjualan!jenis = "Tunai" Then
    Printer.Print Tab(5); "Faktur     :   "; RSPenjualan!faktur
    Printer.Print Tab(5); "Tanggal    :   "; Format(RSPenjualan!tanggal, "DD-MMMM-YYYY")
    Printer.Print Tab(5); "Kasir      :   "; RSKasir!NamaKsr
ElseIf RSPenjualan!jenis = "Kredit" Then
    RSPelanggan.Open "select * From Pelanggan where KodePlg= '" & RSPenjualan!kodePlg & "'", Conn
    Printer.Print Tab(5); "Faktur     :   "; RSPenjualan!faktur
    Printer.Print Tab(5); "Tanggal    :   "; Format(RSPenjualan!tanggal, "DD-MMMM-YYYY")
    Printer.Print Tab(5); "Kasir      :   "; RSKasir!NamaKsr
    Printer.Print Tab(5); "Jenis      :   "; RSPenjualan!jenis
    Printer.Print Tab(5); "Pelanggan  :   "; RSPelanggan!NamaPlg
    Printer.Print Tab(5); "Telepon    :   "; RSPelanggan!teleponPlg
End If
MGrs = String$(33, "-")
Printer.Print Tab(5); MGrs

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
    'Printer berulang-ulang kode,nama,harga,jumlah dan total
    Printer.Print Tab(5); No; Space(2); RSBarang!namabrg
    Printer.Print Tab(10); RKanan(jumlah, "##"); Space(1); "X";
    Printer.Print Tab(15); Format(Harga, "###,###,###");
    Printer.Print Tab(25); RKanan(Hasil, "###,###,###")
    RSDetailJual.MoveNext
Loop

'Printer total harga
If RSPenjualan!jenis = "Tunai" Then
    Printer.Print Tab(5); MGrs
    Printer.Print Tab(5); "Total      :";
    Printer.Print Tab(25); RKanan(RSPenjualan!jmlTotal, "###,###,###");
    Printer.Print Tab(5); "Dibayar    :";
    'Printer dibayar
    Printer.Print Tab(25); RKanan(RSPenjualan!dibayar, "###,###,###");
    Printer.Print Tab(5); MGrs
    Printer.Print Tab(5); "Kembali    :";
    'Printer kembalian
    If RSPenjualan!dibayar = RSPenjualan!jmlTotal Then
        Printer.Print Tab(34); RSPenjualan!dibayar - RSPenjualan!jmlTotal
    Else
        Printer.Print Tab(25); RKanan(RSPenjualan!dibayar - RSPenjualan!jmlTotal, "###,###,###");
    End If
    Printer.Print Tab(5); MGrs
ElseIf RSPenjualan!jenis = "Kredit" Then
    Printer.Print Tab(5); MGrs
    Printer.Print Tab(5); "Total      :";
    Printer.Print Tab(25); RKanan(RSPenjualan!jmlTotal, "###,###,###");
    Printer.Print Tab(5); "Tempo      :";
    Printer.Print Tab(25); RKanan(RSPenjualan!tempo, "###,###,###");
    Printer.Print Tab(5); "Jatuh Tempo:";
    Printer.Print Tab(25); RKanan(RSPenjualan!jatuhtempo, "dd-mmm-yyyy");
    Printer.Print Tab(5); "Uang Muka  :";
    Printer.Print Tab(25); RKanan(RSPenjualan!DP, "###,###,###");
    Printer.Print Tab(5); "Sisa       :";
    Printer.Print Tab(25); RKanan(RSPenjualan!sisa, "###,###,###");
    Printer.Print Tab(5); MGrs
End If
Printer.Print
Printer.Print
Printer.Print
Printer.EndDoc
Conn.Close
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function



