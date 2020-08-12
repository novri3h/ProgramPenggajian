VERSION 5.00
Begin VB.Form Kasir 
   Caption         =   "Data Kasir"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
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
   ScaleHeight     =   1785
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   3360
      TabIndex        =   9
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   2280
      TabIndex        =   8
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "X"
      TabIndex        =   5
      Top             =   840
      Width           =   3200
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   3200
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   1250
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Password"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1000
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1000
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1000
   End
End
Attribute VB_Name = "Kasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Form_Load()
    Call BukaDB
    Text1.MaxLength = 5
    Text2.MaxLength = 25
    Text3.MaxLength = 10
    Kondisiawal
End Sub

Function CariData()
    Call BukaDB
    RSKasir.Open "Select * From Kasir where KodeKsr='" & Text1 & "'", Conn
End Function

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Text3 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
End Sub

Private Sub Kondisiawal()
    KosongkanText
    TidakSiapIsi
    Cmdinput.Caption = "&Input"
    Cmdedit.Caption = "&Edit"
    Cmdhapus.Caption = "&Hapus"
    Cmdtutup.Caption = "&Tutup"
    Cmdinput.Enabled = True
    Cmdedit.Enabled = True
    Cmdhapus.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSKasir
        If Not RSKasir.EOF Then
            Text2 = RSKasir!NamaKsr
            Text3 = RSKasir!PasswordKsr
        End If
    End With
End Sub

Private Sub CmdRefresh_Click()
    Adodc1.Refresh
    DataGrid1.Refresh
    If Cmdinput.Caption = "&Simpan" Then
        Cmdinput.SetFocus
    ElseIf Cmdedit.Caption = "&Simpan" Then
        Cmdedit.SetFocus
    End If
    Adodc1.Refresh
    DataGrid1.Refresh
    Kondisiawal
End Sub


Private Sub AutoNomor()
Call BukaDB
RSKasir.Open ("select * from Kasir Where KodeKsr In(Select Max(KodeKsr)From Kasir)Order By KodeKsr Desc"), Conn
RSKasir.Requery
    Dim Urutan As String * 5
    Dim Hitung As Long
    With RSKasir
        If .EOF Then
            Urutan = "KSR" + "01"
            Text1 = Urutan
        Else
            Hitung = Right(!KodeKsr, 2) + 1
            Urutan = "KSR" + Right("00" & Hitung, 2)
        End If
        Text1 = Urutan
    End With
End Sub

Private Sub CmdInput_click()
    If Cmdinput.Caption = "&Input" Then
        Cmdinput.Caption = "&Simpan"
        Cmdedit.Enabled = False
        Cmdhapus.Enabled = False
        Cmdtutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Call AutoNomor
        Text1.Enabled = False
        Text2.SetFocus
    Else
        If Text1 = "" Or Text2 = "" Or Text3 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Kasir (KodeKsr,NamaKsr,PasswordKsr) values ('" & Text1 & "','" & Text2 & "','" & Text3 & "')"
            Conn.Execute SQLTambah
            Call Kondisiawal
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    If Cmdedit.Caption = "&Edit" Then
        Cmdinput.Enabled = False
        Cmdedit.Caption = "&Simpan"
        Cmdhapus.Enabled = False
        Cmdtutup.Caption = "&Batal"
        SiapIsi
        Text1.SetFocus
    Else
        If Text2 = "" Or Text3 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Kasir Set NamaKsr= '" & Text2 & "', PasswordKsr='" & Text3 & "' where KodeKsr='" & Text1 & "'"
            Conn.Execute SQLEdit
            Call Kondisiawal
        End If
    End If
End Sub

Private Sub CmdHapus_Click()
    If Cmdhapus.Caption = "&Hapus" Then
        Cmdinput.Enabled = False
        Cmdedit.Enabled = False
        Cmdtutup.Caption = "&Batal"
        KosongkanText
        SiapIsi
        Text1.SetFocus
    End If
End Sub

Private Sub CmdTutup_Click()
    Select Case Cmdtutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            Kondisiawal
    End Select
End Sub

Private Sub Text1_Keypress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Len(Text1) < 5 Then
        MsgBox "Kode Harus 5 Digit"
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If

    If Cmdinput.Caption = "&Simpan" Then
        Call CariData
            If Not RSKasir.EOF Then
                TampilkanData
                MsgBox "Kode Kasir Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If Cmdedit.Caption = "&Simpan" Then
        Call CariData
            If Not RSKasir.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Kasir Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If Cmdhapus.Enabled = True Then
        Call CariData
            If Not RSKasir.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Kasir where KodeKsr= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    Kondisiawal
                Else
                    Kondisiawal
                    Cmdhapus.SetFocus
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                Text1.SetFocus
            End If
    End If
End If
End Sub

Private Sub Text2_Keypress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_Keypress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Cmdinput.Enabled = True Then
            Cmdinput.SetFocus
        ElseIf Cmdedit.Enabled = True Then
            Cmdedit.SetFocus
        End If
    End If
End Sub

