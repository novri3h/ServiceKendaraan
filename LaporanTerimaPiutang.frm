VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LaporanTerimaPiutang 
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   1320
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulanan"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   4500
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   720
         Width           =   1500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cetak"
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1250
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Harian"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   4500
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cetak"
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mingguan"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   4500
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   720
         Width           =   1500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cetak"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Awal"
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Akhir"
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1250
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Terima Piutang"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "LaporanTerimaPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
            MsgBox "pilih TanggalTerima yang berbeda"
            Combo3.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{TERIMAPIUTANG.TanggalTerima} in date (" & Combo2 & ") to date (" & Combo3 & ")"
    CR.ReportFileName = App.Path & "\Lap bayar terimapiutang Mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

Private Sub Command3_Click()
    Call BukaDB
    RSPiutang.Open "select * from TERIMAPIUTANG where month(TanggalTerima)='" & Val(Left(Combo4, 2)) & "' and year(TanggalTerima)='" & (Combo5) & "' and keterangan='" & Combo1 & "'", Conn
    If RSPiutang.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    CR.SelectionFormula = "Month({TERIMAPIUTANG.TanggalTerima})=" & Val(Left(Combo4, 2)) & " and Year({terimapiutang.TanggalTerima})=" & Val(Combo5.Text) & " 'and {terimapiutang.keterangan}='" & Combo1 & "'"
    CR.ReportFileName = App.Path & "\MASTER TERIMA PIUTANG.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

Private Sub Command4_Click()
    'CR.SelectionFormula = "{TERIMAPIUTANG.Jenis}='Kredit'"
    CR.ReportFileName = App.Path & "\Lap terimapiutang.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Form_Load()
'On Error Resume Next
Call BukaDB
RSPiutang.Open "Select Distinct keterangan From TERIMAPIUTANG order By 1", Conn
RSPiutang.Requery
Do Until RSPiutang.EOF
    Combo1.AddItem RSPiutang!keterangan
    'Combo2.AddItem Format(RSPiutang!TanggalTerima, "YYYY ,MM, DD")
    'Combo3.AddItem Format(RSPiutang!TanggalTerima, "YYYY ,MM, DD")
    RSPiutang.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGL As New ADODB.Recordset
RSTGL.Open "select distinct month(TanggalTerima) as Bulan from TERIMAPIUTANG", Conn
Do While Not RSTGL.EOF
    Combo4.AddItem RSTGL!bulan & Space(5) & MonthName(RSTGL!bulan)
    RSTGL.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(TanggalTerima)  as Tahun from TERIMAPIUTANG", Conn
Do While Not RSTHN.EOF
    Combo5.AddItem RSTHN!tahun
    RSTHN.MoveNext
Loop
Conn.Close

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
            MsgBox "pilih TanggalTerima yang berbeda"
            Combo3.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{TERIMAPIUTANG.TanggalTerima} in date (" & Combo2 & ") to date (" & Combo3 & ")"
    CR.ReportFileName = App.Path & "\Lap bayar terimapiutang Mingguan.rpt"
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
    'CR.SelectionFormula = "Totext({TERIMAPIUTANG.TanggalTerima})='" & CDate(Combo1) & "'"
    CR.SelectionFormula = "{TERIMAPIUTANG.keterangan}='" & (Combo1) & "'"
    CR.ReportFileName = App.Path & "\Lap piutang.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
'    CR.SelectionFormula = "{terimapiutang.Jenis}='Kredit'"
'    CR.ReportFileName = App.Path & "\Lap terimapiutang.rpt"
'    CR.WindowState = crptMaximized
'    CR.RetrieveDataFiles
'    CR.Action = 1
End Sub


