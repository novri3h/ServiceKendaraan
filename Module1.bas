Attribute VB_Name = "Module1"

Public Conn As New ADODB.Connection
Public RSProfil As ADODB.Recordset
Public RSProfil1 As ADODB.Recordset
Public RSBarang As ADODB.Recordset
Public RSKasir As ADODB.Recordset
Public RSPembelian As ADODB.Recordset
Public RSPendaftaran As ADODB.Recordset

Public RSJasa As ADODB.Recordset
Public RSDetailJasa As ADODB.Recordset
Public RSMekanik As ADODB.Recordset


Public RSDetailBeli As ADODB.Recordset
Public RSReturBeli As ADODB.Recordset
Public RSDetailReturBeli As ADODB.Recordset

Public RSPenjualan As ADODB.Recordset
Public RSDetailJual As ADODB.Recordset

Public RSService As ADODB.Recordset
Public RSDetailService As ADODB.Recordset


Public RSReturJual As ADODB.Recordset
Public RSDetailReturJual As ADODB.Recordset

Public RSTransaksi As ADODB.Recordset
Public RSTransaksi1 As ADODB.Recordset
Public RSPemasok As ADODB.Recordset
Public RSPelanggan As ADODB.Recordset
Public RSHutang As ADODB.Recordset
Public RSPiutang As ADODB.Recordset
Public RSKas As ADODB.Recordset
Public PathData As String

Public Sub BukaDB()
Dim STR As String
Set Conn = New ADODB.Connection
Set RSProfil = New ADODB.Recordset
Set RSProfil1 = New ADODB.Recordset
Set RSBarang = New ADODB.Recordset
Set RSKasir = New ADODB.Recordset
Set RSPendaftaran = New ADODB.Recordset

Set RSJasa = New ADODB.Recordset
Set RSDetailJasa = New ADODB.Recordset
Set RSMekanik = New ADODB.Recordset

Set RSPembelian = New ADODB.Recordset
Set RSDetailBeli = New ADODB.Recordset
Set RSReturBeli = New ADODB.Recordset
Set RSDetailReturBeli = New ADODB.Recordset

Set RSPenjualan = New ADODB.Recordset
Set RSDetailJual = New ADODB.Recordset

Set RSService = New ADODB.Recordset
Set RSDetailService = New ADODB.Recordset

Set RSReturJual = New ADODB.Recordset
Set RSDetailReturJual = New ADODB.Recordset

Set RSTransaksi = New ADODB.Recordset
Set RSTransaksi1 = New ADODB.Recordset
Set RSPemasok = New ADODB.Recordset
Set RSPelanggan = New ADODB.Recordset
Set RSHutang = New ADODB.Recordset
Set RSPiutang = New ADODB.Recordset
Set RSKas = New ADODB.Recordset
PathData = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBRetail.mdb"
Conn.Open PathData
End Sub

