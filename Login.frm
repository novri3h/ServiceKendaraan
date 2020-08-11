VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKodeKsr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3960
      TabIndex        =   6
      Top             =   2280
      Width           =   2000
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   3375
      Begin VB.TextBox TxtPasswordKsr 
         Height          =   350
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   2000
      End
      Begin VB.TextBox TxtNamaKsr 
         Height          =   350
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   2000
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1000
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1000
      End
   End
   Begin VB.Image Image1 
      Height          =   1785
      Left            =   120
      Picture         =   "Login.frx":0000
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   5
      Top             =   2280
      Width           =   1005
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim A As Byte
Dim B As Byte

Private Sub Form_Load()
'batasi jumlah karakter
TxtNamaKsr.MaxLength = 30
TxtPasswordKsr.MaxLength = 10
'TxtPasswordKsr.PasswordChar = "X"
TxtPasswordKsr.Enabled = False
TxtKodeKsr.Enabled = False
End Sub

Private Sub TxtNamaKsr_KeyPress(KeyAscii As Integer)
'ubah karakter jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'jika menekan ESC form ditutup
If KeyAscii = 27 Then Unload Me
'jika menekan enter setelah mengisi nama, maka..
If KeyAscii = 13 Then
    'buka database
    Call BukaDB
    'cari nama kasir yang diketik
    RSKasir.Open "Select NamaKsr from Kasir where NamaKsr ='" & TxtNamaKsr & "'", Conn
    'jika tidak ditemukan, maka
    If RSKasir.EOF Then
        'batasi akses ke nama kasir 3 kali kesempatan
        A = A + 1
        If 1 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 2 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 3 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal" & Chr(13) & _
                    "Kesempatan habis, Ulangi dari awal"
            End
        End If
    Else
        'jika nama kasir benar, maka nama kasir menjadi false
        TxtNamaKsr.Enabled = False
        'password kasir menjadi true dan menjadi fokus kursor
        TxtPasswordKsr.Enabled = True
        TxtPasswordKsr.SetFocus
    End If
End If
End Sub

'coding ini sama dengan nama kasir
Private Sub TxtPasswordKsr_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 27 Then Unload Me
'Dim KodeKasir As String
'Dim NamaKasir As String
If KeyAscii = 13 Then
    Call BukaDB
    RSKasir.Open "Select * from Kasir where NamaKsr ='" & TxtNamaKsr & "' and PassKsr='" & TxtPasswordKsr & "'", Conn
    If RSKasir.EOF Then
        B = B + 1
        If 1 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 2 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 3 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            End
        End If
    Else
        Me.Visible = False
        Menu.Show
        Menu.STBar.Panels(1).Text = RSKasir!kodeksr
        Menu.STBar.Panels(1).Visible = False
        Menu.STBar.Panels(2).Text = Login.TxtNamaKsr
        Menu.STBar.Panels(3).Text = RSKasir!statusksr
        
        If Menu.STBar.Panels(3) = "KASIR" Then
            Menu.mnfile.Enabled = False
            Menu.mnutility.Enabled = False
            Menu.SSTab1.TabEnabled(0) = False
            Menu.SSTab1.TabEnabled(3) = False
            Menu.Image1.Enabled = False
            Menu.Image2.Enabled = False
            Menu.Image3.Enabled = False
            Menu.Image4.Enabled = False
            Menu.Image5.Enabled = False
            Menu.Image13.Enabled = False
            Menu.Image14.Enabled = False
            Menu.Image26.Enabled = False
            Menu.Image27.Enabled = False
            Penjualan.Show
        End If
    End If
End If
End Sub

