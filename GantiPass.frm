VERSION 5.00
Begin VB.Form GantiPass 
   Caption         =   "Ganti Password"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3540
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
   ScaleHeight     =   2370
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Konfirmasi Pwd Baru"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1650
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Password Baru"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1650
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Password Lama"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1650
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1650
   End
End
Attribute VB_Name = "GantiPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BukaDB
    RSKasir.Open "select * from kasir where namaksr='" & Text1 & "'", Conn
    If Not RSKasir.EOF Then
        Text2.SetFocus
    Else
        MsgBox "nama kasir tidak terdaftar"
        Text1.SetFocus
        Text1 = ""
    End If
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BukaDB
    RSKasir.Open "select * from kasir where namaksr='" & Text1 & "' and PassKsr='" & Text2 & "'", Conn
    If Not RSKasir.EOF Then
        Text3.SetFocus
    Else
        MsgBox "password salah "
        Text2.SetFocus
        Text2 = ""
    End If
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text3 = "" Then
        MsgBox "password baru belum dibuat"
        Text3.SetFocus
    Else
        Text4.SetFocus
    End If
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text4 <> Text3 Then
        MsgBox "password konfirmasi tidak sama"
        Text4.SetFocus
        Text4 = ""
    Else
        Pesan = MsgBox("yakin password akan diganti", vbYesNo)
        If Pesan = vbYes Then
            Dim editpwd As String
            editpwd = "update kasir set PassKsr='" & Text4 & "' where namaksr='" & Text1 & "' and PassKsr='" & Text2 & "'"
            Conn.Execute editpwd
            Unload Me
        Else
            Unload Me
        End If
    End If
End If

End Sub

