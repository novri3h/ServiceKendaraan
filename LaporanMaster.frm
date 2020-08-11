VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LaporanMaster 
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2190
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   2190
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   1440
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Pilih Tabel :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "LaporanMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.AddItem "Barang"
List1.AddItem "Kasir"
List1.AddItem "Pelanggan"
List1.AddItem "Pemasok"
End Sub

Private Sub List1_Click()
Select Case List1.ListIndex
    Case 0
        CR.DataFiles(0) = App.Path & "\dbretail.mdb"
        CR.ReportFileName = App.Path & "\Lap Barang.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case 1
        CR.DataFiles(0) = App.Path & "\dbretail.mdb"
        CR.ReportFileName = App.Path & "\Lap Pemakai.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case 2
        CR.DataFiles(0) = App.Path & "\dbretail.mdb"
        CR.ReportFileName = App.Path & "\Lap Pelanggan.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case 3
        CR.DataFiles(0) = App.Path & "\dbretail.mdb"
        CR.ReportFileName = App.Path & "\Lap Pemasok.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
End Select
End Sub
