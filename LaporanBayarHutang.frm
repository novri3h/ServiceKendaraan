VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LaporanBayarHutang 
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   1920
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mingguan"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4500
      Begin VB.CommandButton Command2 
         Caption         =   "Cetak"
         Height          =   375
         Left            =   3000
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Akhir"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1250
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Awal"
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Harian"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4500
      Begin VB.CommandButton Command1 
         Caption         =   "Cetak"
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulanan"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   4500
      Begin VB.CommandButton Command3 
         Caption         =   "Cetak"
         Height          =   375
         Left            =   3000
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1250
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pembayaran Hutang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "LaporanBayarHutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error Resume Next
Call BukaDB
RSHutang.Open "Select Distinct KETERANGAN From bayarHutang order By 1", Conn
RSHutang.Requery
Do Until RSHutang.EOF
    Combo1.AddItem RSHutang!keterangan
    'Combo2.AddItem Format(RSHutang!TanggalBayar, "YYYY ,MM, DD")
    'Combo3.AddItem Format(RSHutang!TanggalBayar, "YYYY ,MM, DD")
    RSHutang.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGL As New ADODB.Recordset
RSTGL.Open "select distinct month(TanggalBayar) as Bulan from bayarHutang", Conn
Do While Not RSTGL.EOF
    Combo4.AddItem RSTGL!bulan & Space(5) & MonthName(RSTGL!bulan)
    RSTGL.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(TanggalBayar)  as Tahun from bayarHutang", Conn
Do While Not RSTHN.EOF
    Combo5.AddItem RSTHN!tahun
    RSTHN.MoveNext
Loop
Conn.Close

End Sub

Private Sub Command2_Click()
    If Combo2 = "" Then
        MsgBox "Tanggal awal kosong", , "Informasi"
        Combo2.SetFocus
        Exit Sub
    Else
        If Combo3 < Combo2 Or Combo2 > Combo3 Then
            MsgBox "Tanggal terbalik"
            Combo3.SetFocus
            Exit Sub
        ElseIf Combo3 = Combo2 Then
            MsgBox "pilih TanggalBayar yang berbeda"
            Combo3.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{bayarHutang.TanggalBayar} in date (" & Combo2 & ") to date (" & Combo3 & ")"
    CR.ReportFileName = App.Path & "\Lap bayar hutang Mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

Private Sub Command3_Click()
    Call BukaDB
    RSHutang.Open "select * from bayarHutang where month(TanggalBayar)='" & Val(Left(Combo4, 2)) & "' and year(TanggalBayar)='" & (Combo5) & "' AND KETERANGAN='" & Combo1 & "'", Conn
    If RSHutang.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    CR.SelectionFormula = "Month({bayarHutang.TanggalBayar})=" & Val(Left(Combo4, 2)) & " and Year({BAYARHutang.TanggalBayar})=" & Val(Combo5.Text) & " AND {BAYARHUTANG.KETERANGAN}='" & Combo1 & "'"
    CR.ReportFileName = App.Path & "\MASTER BAYAR HUTANG.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub



'Lap Harian
Private Sub Combo1_Click()
End Sub

'Lap Mingguan (Tgl Antara)
Private Sub Combo3_Click()
    If Combo2 = "" Then
        MsgBox "Tanggal awal kosong", , "Informasi"
        Combo2.SetFocus
        Exit Sub
    Else
        If Combo3 < Combo2 Or Combo2 > Combo3 Then
            MsgBox "Tanggal terbalik"
            Combo3.SetFocus
            Exit Sub
        ElseIf Combo3 = Combo2 Then
            MsgBox "pilih TanggalBayar yang berbeda"
            Combo3.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{bayarHutang.TanggalBayar} in date (" & Combo2 & ") to date (" & Combo3 & ")"
    CR.ReportFileName = App.Path & "\Lap bayar hutang Mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'Lap Bulanan
Private Sub Combo5_Click()
End Sub

Private Sub Command1_Click()
If Combo1 = "" Then
    MsgBox "Pilih tanggal.."
    Combo1.SetFocus
    Exit Sub
End If
    CR.SelectionFormula = "Totext({bayarHutang.TanggalBayar})='" & CDate(Combo1) & "'"
    CR.ReportFileName = App.Path & "\Lap bayar hutang harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
'    CR.SelectionFormula = "{Hutang.Jenis}='Kredit'"
'    CR.ReportFileName = App.Path & "\Lap hutang.rpt"
'    CR.WindowState = crptMaximized
'    CR.RetrieveDataFiles
'    CR.Action = 1
End Sub

