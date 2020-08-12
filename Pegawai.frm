VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pegawai 
   Caption         =   "Data Pegawai"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "NIP"
         Caption         =   "NIP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NamaPgw"
         Caption         =   "Nama Pegawai"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Bagian"
         Caption         =   "Bagian"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995,024
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmdrefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton Cmdtutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton Cmdhapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton Cmdedit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   1000
   End
   Begin VB.CommandButton Cmdinput 
      Caption         =   "&Input"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   1250
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   4260
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1250
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bagian"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "Pegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mvBookMark As Variant

Private Sub Form_Activate()
Call BukaDB
Conn.CursorLocation = adUseClient
RSPegawai.Open "select * from Pegawai", Conn
With RSPegawai
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
End With
Set DataGrid1.DataSource = RSPegawai.DataSource
End Sub

Sub Form_Load()
Text1.MaxLength = 9
Text2.MaxLength = 30
Text3.MaxLength = 20
Kondisiawal
End Sub

Function CariData()
    Call BukaDB
    RSPegawai.Open "Select * From Pegawai where NIP='" & Text1 & "'", Conn
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
    With RSPegawai
        If Not RSPegawai.EOF Then
            Text2 = RSPegawai!NamaPgw
            Text3 = RSPegawai!Bagian
        End If
    End With
End Sub

Private Sub CmdRefresh_Click()
    If Cmdinput.Caption = "&Simpan" Then
        Cmdinput.SetFocus
    ElseIf Cmdedit.Caption = "&Simpan" Then
        Cmdedit.SetFocus
    End If
    Call Kondisiawal
    Form_Activate
End Sub

Private Sub CmdInput_click()
    If Cmdinput.Caption = "&Input" Then
        Cmdinput.Caption = "&Simpan"
        Cmdedit.Enabled = False
        Cmdhapus.Enabled = False
        Cmdtutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Text1.SetFocus
    Else
        If Text1 = "" Or Text2 = "" Or Text3 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Pegawai (NIP,NamaPgw,Bagian) values ('" & Text1 & "','" & Text2 & "','" & Text3 & "')"
            Conn.Execute SQLTambah
            Cmdrefresh.SetFocus
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
            SQLEdit = "Update Pegawai Set NamaPgw= '" & Text2 & "', Bagian='" & Text3 & "' where NIP='" & Text1 & "'"
            Conn.Execute SQLEdit
            Cmdrefresh.SetFocus
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
    If Len(Text1) < 9 Then
        MsgBox "Kode Harus 9 Digit"
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If

    If Cmdinput.Caption = "&Simpan" Then
        Call CariData
            If Not RSPegawai.EOF Then
                TampilkanData
                MsgBox "Kode Pegawai Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If Cmdedit.Caption = "&Simpan" Then
        Call CariData
            If Not RSPegawai.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Pegawai Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If Cmdhapus.Enabled = True Then
        Call CariData
            If Not RSPegawai.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Pegawai where NIP= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    Kondisiawal
                    Cmdrefresh.SetFocus
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
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
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

