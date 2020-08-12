VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Perkiraan 
   Caption         =   "Data Perkiraan"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   3810
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3625
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "KodePrk"
         Caption         =   "Kode"
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
         DataField       =   "NamaPrk"
         Caption         =   "Nama"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   1250
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Width           =   4260
   End
   Begin VB.CommandButton Cmdinput 
      Caption         =   "&Input"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton Cmdedit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton Cmdhapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton Cmdtutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton Cmdrefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   960
      Width           =   1000
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
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1005
   End
End
Attribute VB_Name = "Perkiraan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mvBookMark As Variant

Private Sub Form_Activate()
Call BukaDB
Conn.CursorLocation = adUseClient
RSPerkiraan.Open "select * from Perkiraan", Conn
With RSPerkiraan
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
End With
Set DataGrid1.DataSource = RSPerkiraan.DataSource
End Sub

Sub Form_Load()
Text1.MaxLength = 6
Text2.MaxLength = 30
Kondisiawal
End Sub

Function CariData()
    Call BukaDB
    RSPerkiraan.Open "Select * From Perkiraan where KodePrk='" & Text1 & "'", Conn
End Function

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    
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
    With RSPerkiraan
        If Not RSPerkiraan.EOF Then
            Text2 = RSPerkiraan!NamaPrk
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
        If Text1 = "" Or Text2 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Perkiraan (KodePrk,NamaPrk) values ('" & Text1 & "','" & Text2 & "')"
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
        If Text2 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Perkiraan Set NamaPrk= '" & Text2 & "' where KodePrk='" & Text1 & "'"
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
    If Len(Text1) < 3 Then
        MsgBox "Kode Harus 3 Digit"
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If

    If Cmdinput.Caption = "&Simpan" Then
        Call CariData
            If Not RSPerkiraan.EOF Then
                TampilkanData
                MsgBox "Kode Perkiraan Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If Cmdedit.Caption = "&Simpan" Then
        Call CariData
            If Not RSPerkiraan.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Perkiraan Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If Cmdhapus.Enabled = True Then
        Call CariData
            If Not RSPerkiraan.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Perkiraan where kodePrk= '" & Text1 & "'"
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
    If KeyAscii = 13 Then
        If Cmdinput.Enabled = True Then
            Cmdinput.SetFocus
        ElseIf Cmdedit.Enabled = True Then
            Cmdedit.SetFocus
        End If
    End If
End Sub

