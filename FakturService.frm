VERSION 5.00
Begin VB.Form FakturService 
   BackColor       =   &H80000009&
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5400
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
   ScaleHeight     =   7725
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FakturService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
    Call CetakPrinter
End If
End Sub

Function CetakPrinter()

Call BukaDB
RSService.Open "select * from Service Where Faktur In(Select Max(Faktur)From Service)Order By Faktur Desc", Conn
Dim JmlHarga, JmlJual, JmlHasil As Double
Dim MGrs As String
Printer.Font = "Courier New"
Printer.Print
Printer.Print
Printer.CurrentX = 0
Printer.CurrentY = 0
RSKasir.Open "select * From Kasir where KodeKsr= '" & RSService!kodeksr & "'", Conn
Printer.Print Tab(5); "Faktur     :   "; RSService!faktur
Printer.Print Tab(5); "Tanggal    :   "; Format(RSService!tanggal, "DD-MMMM-YYYY")
Printer.Print Tab(5); "Kasir      :   "; RSKasir!NamaKsr
Printer.Print Tab(5); "No Polisi  :   "; RSService!nopol
MGrs = String$(40, "-")
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "JASA / PELAYANAN"

RSDetailJasa.Open "select * from DetailJasa Where Faktur='" & RSService!faktur & "'", Conn
RSDetailJasa.MoveFirst
No = 0
Do While Not RSDetailJasa.EOF
    No = No + 1
    Harga = RSDetailJasa!Harga
    Printer.Print Tab(5); No; Space(2); RSDetailJasa!nama_jasa;
    Printer.Print Tab(30); RKanan(Harga, "###,###,###");
    RSDetailJasa.MoveNext
Loop
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Total Biaya Jasa   :   ";
Printer.Print Tab(30); RKanan(RSService!biayajasa, "###,###,###");

Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "SPARE PART"

RSDetailService.Open "select * from DetailService Where Faktur='" & RSService!faktur & "'", Conn
RSDetailService.MoveFirst
No = 0
Do While Not RSDetailService.EOF
    No = No + 1
    Harga = RSDetailService!HargaJual
    jumlah = RSDetailService!JmlJual
    Hasil = Harga * jumlah

    Printer.Print Tab(5); No; Space(2); RSDetailService!NamaBarang;
    Printer.Print Tab(10); RKanan(jumlah, "##"); Space(1); "X";
    Printer.Print Tab(15); RKanan(Harga, "###,###,###");
    Printer.Print Tab(30); RKanan(Hasil, "###,###,###")
    RSDetailService.MoveNext
Loop
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Total Biaya Barang : ";
Printer.Print Tab(30); RKanan(RSService!biayabarang, "###,###,###");

Printer.Print Tab(5); MGrs
If RSService!diskon = 0 Then
    Printer.Print Tab(5); "Diskon    : ";
    Printer.Print Tab(39); 0
Else
    Printer.Print Tab(5); "Diskon    : ";
    Printer.Print Tab(30); RKanan(RSService!diskon, "###,###,###");
End If

Printer.Print Tab(5); "Total      :";
Printer.Print Tab(30); RKanan(RSService!jmlTotal, "###,###,###");
Printer.Print Tab(5); "Dibayar    :";

Printer.Print Tab(30); RKanan(RSService!dibayar, "###,###,###");
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Kembali    :";
If RSService!dibayar = RSService!jmlTotal Then
    Printer.Print Tab(39); RSService!dibayar - RSService!jmlTotal
Else
    Printer.Print Tab(30); RKanan(RSService!dibayar - RSService!jmlTotal, "###,###,###");
End If
Printer.Print Tab(5); MGrs

Printer.Print
Conn.Close
Printer.EndDoc
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

