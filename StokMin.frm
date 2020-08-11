VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form StokMin 
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5490
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
   ScaleHeight     =   2190
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cetak"
      Height          =   400
      Left            =   4080
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak"
      Height          =   400
      Left            =   4080
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin Crystal.CrystalReport CR 
      Left            =   840
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Laporan Stok Barang"
      BeginProperty Font 
         Name            =   "Century"
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
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stok Barang Lebih Dari"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stok Barang Kurang Dari"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "StokMin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call BukaDB
RSBarang.Open "select * from barang where val(jumlahbrg)<='" & Val(Combo1) & "'", Conn
If RSBarang.EOF Then
    MsgBox "Data tidak ditemukan"
    Exit Sub
Else
    CR.SelectionFormula = "({Barang.jumlahbrg})<=" & Combo1 & ""
    CR.ReportFileName = App.Path & "\Lap Barang.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End If
End Sub

Private Sub Command2_Click()
If Combo2 = "" Then
    MsgBox "Pilih jumlah"
    Combo2.SetFocus
    Exit Sub
Else
    CR.SelectionFormula = "({Barang.jumlahbrg})>=" & Combo2 & ""
    CR.ReportFileName = App.Path & "\Lap Barang.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End If
End Sub

Private Sub Form_Load()
For i = 5 To 50 Step 5
    Combo1.AddItem i
    Combo2.AddItem i
Next i
End Sub
    
