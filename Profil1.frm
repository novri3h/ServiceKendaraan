VERSION 5.00
Begin VB.Form Profil1 
   Caption         =   "Profil Perusahaan"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan Data"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Profil1.frx":0000
      Top             =   120
      Width           =   6500
   End
End
Attribute VB_Name = "Profil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call BukaDB
RSProfil1.Open "select * from Profil1", Conn
If RSProfil1.EOF Then
    Dim Simpan As String
    Simpan = "insert into Profil1(Data) values ('" & Text1 & "')"
    Conn.Execute Simpan
    MsgBox "Data telah berhasil disimpan"
    Unload Me
Else
    Dim edit As String
    edit = "update Profil1 set Data='" & Text1 & "'"
    Conn.Execute edit
    MsgBox "Data telah berhasil diubah"
    Unload Me
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Call BukaDB
RSProfil1.Open "Profil1", Conn
Text1 = RSProfil1!Data
End Sub


Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then Command1.SetFocus
End Sub

