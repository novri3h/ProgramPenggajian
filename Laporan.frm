VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Laporan 
   Caption         =   "Laporan"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3615
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
   ScaleHeight     =   3510
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Laporan Akumulasi"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   3360
      Begin VB.CommandButton Command2 
         Caption         =   "Print"
         Height          =   615
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   1250
      End
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   1000
      End
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   750
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Laporan Per Pegawai"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3360
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   615
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   1250
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   1000
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   750
      End
   End
End
Attribute VB_Name = "Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Private Sub Form_Load()
''buatlah looping untuk bulan dari 1-12
''dan tahun mulai tahun 2001 s/d 2020
'For i = 1 To 12
'    Combo1.AddItem i
'    Combo3.AddItem i
'Next i
'Combo1.Text = Month(Date)
'Combo2.Text = Year(Date)
'
'For i = 1 To 10
'    Combo2.AddItem 2000 + i
'    Combo4.AddItem 2000 + i
'Next i
'
'Combo3.Text = Month(Date)
'Combo4.Text = Year(Date)
'End Sub


Private Sub Form_Load()
'On Error Resume Next
'Call BukaDB
'RSGaji.Open "Select Distinct Tanggal From gaji order By 1", Conn
'RSGaji.Requery
'Do Until RSGaji.EOF
'    Combo1.AddItem Format(RSGaji!Tanggal, "DD-MMM-YYYY")
'    Combo2.AddItem Format(RSGaji!Tanggal, "YYYY ,MM, DD")
'    Combo3.AddItem Format(RSGaji!Tanggal, "YYYY ,MM, DD")
'    RSGaji.MoveNext
'Loop
'Conn.Close

Call BukaDB
Dim RSTGL As New ADODB.Recordset
RSTGL.Open "select distinct month(Tanggal) as Bulan from gaji", Conn
Do While Not RSTGL.EOF
    Combo1.AddItem RSTGL!Bulan & Space(5) & MonthName(RSTGL!Bulan)
    Combo3.AddItem RSTGL!Bulan & Space(5) & MonthName(RSTGL!Bulan)
    RSTGL.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(Tanggal)  as Tahun from gaji", Conn
Do While Not RSTHN.EOF
    Combo2.AddItem RSTHN!Tahun
    Combo4.AddItem RSTHN!Tahun
    RSTHN.MoveNext
Loop
Conn.Close

End Sub


Private Sub Command1_Click()
    'jika bulan dan tahun masih kosong'
    'tampilkan pesan...
    If Combo1 = "" Or Combo2 = "" Then
        MsgBox "pilih bulan dan tahun...!"
        Combo1.SetFocus
        Exit Sub
    Else
        'buka database
        Call BukaDB
        'cari data yang tanggal dan bulannya dipilih di Combo1 dan Combo2
        RSGaji.Open "select * from Gaji where month(tanggal)='" & Val(Combo1) & "' and year(tanggal)='" & (Combo2) & "'", Conn
        'jika tidak cocok, munculkan pesan
        If RSGaji.EOF Then
            MsgBox "Data tidak ditemukan"
            Exit Sub
            Combo1.SetFocus
        End If
        'jika ditemukan panggil file laporan yang
        'datanya bulannya=Combo1 dan tahunnya= Combo2
        CR.SelectionFormula = "Month({Gaji.Tanggal})=" & Val(Combo1.Text) & " and Year({Gaji.Tanggal})=" & Val(Combo2.Text)
        CR.ReportFileName = App.Path & "\Lap Gaji.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    End If
End Sub

Private Sub Command2_Click()
    If Combo3 = "" Or Combo4 = "" Then
        MsgBox "pilih bulan dan tahun...!"
        Combo1.SetFocus
        Exit Sub
    Else
        'buka database
        Call BukaDB
        'cari data yang tanggal dan bulannya dipilih di Combo1 dan combo4
        RSGaji.Open "select * from Gaji where month(tanggal)='" & Val(Combo3) & "' and year(tanggal)='" & (Combo4) & "'", Conn
        'jika tidak cocok, munculkan pesan
        If RSGaji.EOF Then
            MsgBox "Data tidak ditemukan"
            Exit Sub
            Combo1.SetFocus
        End If
        'jika ditemukan panggil file laporan yang
        'datanya bulannya=Combo1 dan tahunnya= combo4
        CR.SelectionFormula = "Month({Gaji.Tanggal})=" & Val(Combo3.Text) & " and Year({Gaji.Tanggal})=" & Val(Combo4.Text)
        CR.ReportFileName = App.Path & "\Lap Gaji1.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    End If
End Sub

