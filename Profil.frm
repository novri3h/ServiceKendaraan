VERSION 5.00
Begin VB.Form Profil 
   Caption         =   "Data Perusahaan"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   6500
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   6500
   End
   Begin VB.TextBox Text3 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   6500
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   6500
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   6500
   End
   Begin VB.TextBox Text6 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   6500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan Data"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama "
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pemilik"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat1"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat2"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Email"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1500
   End
End
Attribute VB_Name = "Profil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call BukaDB
RSProfil.Open "select * from profil", Conn
If RSProfil.EOF Then
    Dim Simpan As String
    Simpan = "insert into profil(nama_perusahaan,pemilik,alamat1,alamat2,telepon,email) values " & _
    "('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "')"
    Conn.Execute Simpan
    MsgBox "Data telah berhasil disimpan"
    Unload Me
Else
    Dim edit As String
    edit = "update profil set Nama_Perusahaan='" & Text1 & "', " & _
    "pemilik='" & Text2 & "', " & _
    "alamat1='" & Text3 & "', " & _
    "alamat2='" & Text4 & "', " & _
    "telepon='" & Text5 & "', " & _
    "email='" & Text6 & "'"
    Conn.Execute edit
    MsgBox "Data telah berhasil diubah"
    Unload Me
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Call BukaDB
RSProfil.Open "profil", Conn
Text1 = RSProfil!Nama_Perusahaan
Text2 = RSProfil!pemilik
Text3 = RSProfil!alamat1
Text4 = RSProfil!alamat2
Text5 = RSProfil!telepon
Text6 = RSProfil!email
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then Text4.SetFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then Text5.SetFocus
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text6.SetFocus
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

