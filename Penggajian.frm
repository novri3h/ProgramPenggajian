VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Penggajian 
   Caption         =   "Pengolahan Data Penggajian"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
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
   ScaleHeight     =   5475
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   2820
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   850
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batalkan Semua"
      Height          =   350
      Left            =   960
      TabIndex        =   2
      Top             =   4320
      Width           =   1500
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   2520
      TabIndex        =   3
      Top             =   4320
      Width           =   850
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2760
      Top             =   5040
   End
   Begin VB.ListBox List1 
      Height          =   3885
      Left            =   6240
      TabIndex        =   0
      Top             =   1320
      Width           =   2500
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   345
      Left            =   120
      Top             =   4800
      Visible         =   0   'False
      Width           =   2350
      _ExtentX        =   4154
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Transaksi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Penggajian.frx":0000
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5106
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Nomor"
         Caption         =   "Nomor"
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
         DataField       =   "Kode"
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
      BeginProperty Column02 
         DataField       =   "Nama"
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
      BeginProperty Column03 
         DataField       =   "Jumlah"
         Caption         =   "Jumlah"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1500,095
         EndProperty
      EndProperty
   End
   Begin VB.Label LblKodeKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2640
      TabIndex        =   25
      Top             =   4680
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label LblJam 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1320
      TabIndex        =   24
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label LblNamaKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7200
      TabIndex        =   23
      Top             =   120
      Width           =   1480
   End
   Begin VB.Label LblTanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1320
      TabIndex        =   22
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label LblNomorSlp 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1320
      TabIndex        =   21
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label LblPendapatan 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4680
      TabIndex        =   20
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Label LblPotongan 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4680
      TabIndex        =   19
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4680
      TabIndex        =   18
      Top             =   5040
      Width           =   1500
   End
   Begin VB.Label LblNamaPgw 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3720
      TabIndex        =   17
      Top             =   480
      Width           =   4965
   End
   Begin VB.Label LblBagian 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3720
      TabIndex        =   16
      Top             =   840
      Width           =   4965
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bagian"
      Height          =   315
      Left            =   2760
      TabIndex        =   14
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   315
      Left            =   2760
      TabIndex        =   13
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NIP"
      Height          =   315
      Left            =   2760
      TabIndex        =   12
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Gaji Bersih"
      Height          =   315
      Left            =   3600
      TabIndex        =   10
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Potongan"
      Height          =   315
      Left            =   3600
      TabIndex        =   9
      Top             =   4680
      Width           =   1005
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pendapatan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      Top             =   4320
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Slip"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1250
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1250
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kasir"
      Height          =   315
      Left            =   6600
      TabIndex        =   5
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jam"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1250
   End
End
Attribute VB_Name = "Penggajian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'tampilkan waktu
Private Sub Timer1_Timer()
LblJam = Time$
End Sub

Private Sub Form_Activate()
DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOGaji.mdb"
DT.RecordSource = "Transaksi"
Set DataGrid1.DataSource = DT
DataGrid1.Refresh

'jika kode kasir tidak terdeteksi tampilkan pesan ...lalu tampilkan form login
If LblKodeKsr = "" Then
    MsgBox "Kasir tidak terdeteksi"
    Login.Show
    Exit Sub
End If
'kode dan nama kasir diambil dari login
LblKodeKsr = Login.TxtKodeKsr
LblNamaKsr = Login.TxtNamaKsr
'memanggil nomor slip gaji otomatis
Call Auto
'memanggil prosedur agar tabel transaksi dikosongkan
Call Tabel_Kosong
DT.Recordset.MoveFirst
Combo1.SetFocus
If DataGrid1.Columns(1) <> vbNullString Then DataGrid1.Col = 3
End Sub

Private Sub Form_Load()
'buka database
Call BukaDB
'buka tabel perkiraan
RSPerkiraan.Open "select * from Perkiraan", Conn
RSPerkiraan.Requery
List1.Clear
'tampilkan data kode dan nama perkiraan di list
Do Until RSPerkiraan.EOF
    List1.AddItem RSPerkiraan!KodePrk & Space(2) & RSPerkiraan!NamaPrk
    RSPerkiraan.MoveNext
Loop
'buka tabel pegawai
RSPegawai.Open "select * from Pegawai", Conn
RSPegawai.Requery
Combo1.Clear
'tampilkan NIP di combo
Do Until RSPegawai.EOF
    Combo1.AddItem RSPegawai!NIP & Space(5) & RSPegawai!NamaPgw
    RSPegawai.MoveNext
Loop

LblKodeKsr = Login.TxtKodeKsr
LblNamaKsr = Login.TxtNamaKsr
'aktifkan tanggal hari ini
LblTanggal = Format(Date, "dd-mm-yyyy")
End Sub

Private Sub Auto()
Call BukaDB
'baca nomor slip gaji terakhir
RSGaji.Open "select * from Gaji Where NomorSlp In(Select Max(NomorSlp)From Gaji)Order By NomorSlp Desc", Conn
RSGaji.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSGaji
        If .EOF Then
            Urutan = Format(Date, "yymmdd") + "0001"
            LblNomorSlp = Urutan
        Else
            If Left(!NomorSlp, 6) <> Format(Date, "yymmdd") Then
                Urutan = Format(Date, "yymmdd") + "0001"
            Else
                Hitung = (!NomorSlp) + 1
                Urutan = Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
            End If
        End If
        LblNomorSlp = Urutan
    End With
End Sub

'jika dalam tabel transaksi masih ada data
'hapus data tersebut
Function Tabel_Kosong()
If Not DT.Recordset.RecordCount = 0 Then
    DT.Recordset.MoveFirst
    Do While Not DT.Recordset.EOF
        DT.Recordset.Delete
        DT.Recordset.MoveNext
    Loop
End If
'tampilkan 1 nomor transaksi dalam Grid
For i = 1 To 1
    DT.Recordset.AddNew
    DT.Recordset!Nomor = i
    DT.Recordset.Update
Next i
DataGrid1.Col = 1
End Function

'menambah baris transaksi setelah baris diatasnya diisi
Function Tambah_Baris()
For i = DT.Recordset.RecordCount To DT.Recordset.RecordCount
    DT.Recordset.AddNew
    DT.Recordset!Nomor = i + 1
    DT.Recordset.Update
Next i
End Function

Private Sub Combo1_Keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Combo1 = "" Then
        MsgBox "NIP Kosong"
        Exit Sub
        Combo1.SetFocus
    Else
        'cari data gaji berdasarkan NIP dan Bulannya bulan sekarang
        Call BukaDB
        RSGaji.Open "Select * from Gaji where Nip='" & Combo1 & "' and cdate(month(tanggal))='" & CDate(Month(LblTanggal)) & "'", Conn
        'jika ditemukan tampilkan pesan
        If Not RSGaji.EOF Then
            MsgBox "Nip     : " & Combo1 & " " & Chr(13) & _
            "Nama : " & LblNamaPgw & " " & Chr(13) & _
            "Bulan ini telah menerima gaji"
            Combo1.SetFocus
            Exit Sub
        End If
            DataGrid1.SetFocus
        End If
    End If

End Sub

Private Sub Combo1_Click()
Call BukaDB
'cari pegawai berdasarkan NIP
RSPegawai.Open "Select * from pegawai where nip='" & Left(Combo1, 9) & "'", Conn
'jika tidak ditemukan munculkan pesan
If RSPegawai.EOF Then
    MsgBox "Nip tidak terdaftar"
    Combo1.SetFocus
Else
    LblNamaPgw = RSPegawai!NamaPgw
    LblBagian = RSPegawai!Bagian

    RSGaji.Open "select * from gaji where nip='" & Left(Combo1, 9) & "' and month(tanggal)='" & Month(LblTanggal) & "' and year(tanggal)='" & Year(LblTanggal) & "'", Conn
    If Not RSGaji.EOF Then
        MsgBox "nip tsb sudah gajian"
        LblNamaPgw = ""
        LblBagian = ""
        Combo1.SetFocus
        Exit Sub
    Else
    'jika ditemukan tampilkan datanya
    Combo1.SetFocus
    End If
End If
End Sub

'mencari data perkiraan berdasarkan kode yang ada di list
Private Sub list1_keyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If DataGrid1.SelText <> Left(List1, 3) Then
        DataGrid1.SelText = Left(List1, 3)
        Call BukaDB
        RSPerkiraan.Open "Select * from perkiraan where KodePrk='" & Left(List1, 3) & "'", Conn
        RSPerkiraan.Requery
        If Not RSPerkiraan.EOF Then
            'jika kode perkiraannya 101, maka pajaknya lansgung terisi 10%
            'dari total pendapatan
            If Left(List1, 3) = "101" Then
                DT.Recordset!Kode = RSPerkiraan!KodePrk
                DT.Recordset!Nama = RSPerkiraan!NamaPrk
                DT.Recordset!Jumlah = LblPendapatan * 0.1
                DT.Recordset.Update
                Call Tambah_Baris
                DT.Recordset.MoveLast
                LblPendapatan = Format(GajiKotor, "###,###,###")
                LblPotongan = Format(Potongan, "###,###,###")
                LblTotal = Format(GajiKotor - Potongan, "###,###,###")
                DataGrid1.SetFocus
                DataGrid1.Col = 1
            Else
                'jika kodenya bukan 101, maka jumlahnya harus diisi manual
                DT.Recordset!Kode = RSPerkiraan!KodePrk
                DT.Recordset!Nama = RSPerkiraan!NamaPrk
                DT.Recordset.Update
                DataGrid1.SetFocus
                DataGrid1.Col = 3
                DataGrid1.Refresh
            End If
        End If
    End If
End If
End Sub

Private Sub DataGrid1_Keypress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
If DataGrid1.Col = 1 Then
    'cari data perkiraan berdasarkan kode supaya muncul nama perkiraannya
    Call BukaDB
    RSPerkiraan.Open "Select * from perkiraan where kodeprk= '" & DT.Recordset!Kode & "'", Conn
    RSPerkiraan.Requery
    If RSPerkiraan.EOF Then
        MsgBox "Kode Perkiraan tidak terdaftar" & Chr(13) & _
        "Lihat daftar di kanan " & Chr(13) & _
        "Pilih datanya, lalu tekan Enter"
        List1.SetFocus
        Exit Sub
    End If
    'jika kode ditemukan tampilkan nama perkirannya
    DT.Recordset!Kode = RSPerkiraan!KodePrk
    DT.Recordset!Nama = RSPerkiraan!NamaPrk
    DataGrid1.SetFocus
    DataGrid1.Col = 3
    Exit Sub
End If

'jika jumlah nominal elah diisi, tampilkan total pendapatan
'tampilkan total potongan dan gaji bersihnya
If DataGrid1.Col = 3 Then
    DT.Recordset!Jumlah = DT.Recordset!Jumlah
    DT.Recordset.Update
    Call Tambah_Baris
    DT.Recordset.MoveNext
    DataGrid1.SetFocus
    DataGrid1.Col = 1
    DT.Recordset.MoveLast
    LblPendapatan = Format(GajiKotor, "###,###,###")
    LblPotongan = Format(Potongan, "###,###,###")
    LblTotal = Format(GajiKotor - Potongan, "###,###,###")
End If
End Sub

'encari total pendapatan
Function GajiKotor()
    Set RSGajiKotor = New ADODB.Recordset
    RSGajiKotor.Open "Select sum(Jumlah) as TTLPendapatan from Transaksi where left(kode,1)=0", Conn
    GajiKotor = RSGajiKotor!TTLPendapatan
End Function

'mencari total potongan
Function Potongan()
    Set RSPotongan = New ADODB.Recordset
    RSPotongan.Open "Select sum(Jumlah) as TTLPotongan from Transaksi where left(kode,1)=1", Conn
    Potongan = RSPotongan!TTLPotongan
End Function

Private Sub Bersihkan()
LblNamaPgw = ""
LblBagian = ""
LblPendapatan = ""
LblPotongan = ""
LblTotal = ""
End Sub

Private Sub CmdSimpan_Keypress(KeyAscii As Integer)
If KeyAscii = 27 Then DataGrid1.SetFocus
End Sub

Private Sub CmdSimpan_Click()
If Combo1 = "" Then
    MsgBox "NIP karyawan belum diisi"
    Combo1.SetFocus
    Exit Sub
ElseIf LblPendapatan = "" Then
    MsgBox "Data belum lengkap"
    Exit Sub
Else
    Pesan = MsgBox("Data sudah benar..?", vbYesNo, "Konfirmasi")
    'simpan data ke tabel gaji pada bagian header dan footer
    '(disimpan hanya sekali saja)
    If Pesan = vbYes Then
        Dim SimpanGaji As String
        SimpanGaji = "insert into gaji (NomorSlp,tanggal,jam,nip,pendapatan,potongan,gajibersih,kodeksr) values " & _
        "('" & LblNomorSlp & "','" & CDate(LblTanggal) & "','" & LblJam & "', " & _
        "'" & Left(Combo1, 9) & "','" & LblPendapatan & "','" & LblPotongan & "', " & _
        "'" & LblTotal & "','" & LblKodeKsr & "')"
        Conn.Execute (SimpanGaji)
        
        'simpan data berulang-ulang ke tabel detailgaji
        'yg disimpan adalah nomor slip, kode perkiraan dan jumlah nominalnya
        DT.Recordset.MoveFirst
        Do Until DT.Recordset.EOF
            If DT.Recordset!Kode <> vbNullString Then
                Dim simpandetail As String
                simpandetail = "insert into detailgaji(NomorSlp,KodePrk,Jumlah) values " & _
                "('" & LblNomorSlp & "','" & DT.Recordset!Kode & "','" & DT.Recordset!Jumlah & "')"
                Conn.Execute (simpandetail)
            End If
            DT.Recordset.MoveNext
        Loop
        Bersihkan
        Form_Activate
        Call CetakLayar
    Else
        DataGrid1.SetFocus
    End If
End If
End Sub

Private Sub CmdBatal_Click()
Bersihkan
Form_Activate
End Sub

Private Sub CmdTutup_Click()
CmdBatal_Click
Unload Me
End Sub


Sub CetakLayar()
Tampilkan.Show
Dim MGrs As String
Tampilkan.Font = "Courier New"
'cari data gaji dengan nomor slip terakhir
Call BukaDB
RSGaji.Open "select * from Gaji Where NomorSlp In(Select Max(NomorSlp)From Gaji)Order By NomorSlp Desc", Conn
'cari data kasir yang kodenya ada di tabel gaji
RSKasir.Open "select * from kasir where kodeksr='" & RSGaji!KodeKsr & "'", Conn
RSKasir.Requery
'cari data pegawai yang nip-nya ada di tabel gaji
RSPegawai.Open "select * from pegawai where NIP='" & RSGaji!NIP & "'", Conn
RSPegawai.Requery
Tampilkan.Print
Tampilkan.FontBold = True
Tampilkan.Print
Tampilkan.FontBold = False
'cetak nomor slip,tanggal,nip dan seterusnya
Tampilkan.Print Tab(5); "Nomor Slip :  "; RSGaji!NomorSlp
Tampilkan.Print Tab(5); "Tanggal    :  "; Format(RSGaji!Tanggal, "DD-MMMM-YYYY")
Tampilkan.Print Tab(5); "NIP        :  "; RSGaji!NIP
Tampilkan.Print Tab(5); "Nama       :  "; RSPegawai!NamaPgw
Tampilkan.Print Tab(5); "Kasir      :  "; RSKasir!NamaKsr
MGrs = String$(33, "-")
'cetak isi tabel detailgaji berdasarkan nomor slip terakhir
Tampilkan.Print Tab(5); MGrs
RSDetail.Open "select * from DetailGaji Where NomorSlp In(Select max(NomorSlp)From DetailGaji)", Conn
RSDetail.Requery
'If Not RSDetail.EOF Then
RSDetail.MoveFirst
Do While Not RSDetail.EOF
    'Call BukaDB
    Set RSPerkiraan = New ADODB.Recordset
    RSPerkiraan.Open "select * from Perkiraan where KodePrk='" & RSDetail!KodePrk & "'", Conn
    Tampilkan.Print Tab(5); RSPerkiraan!NamaPrk;
    If Left(RSDetail!KodePrk, 1) = "0" Then
        'jika satu digit awal kode perkirannya 0 maka beri tanda plus (+)
        Tampilkan.Print Tab(25); RKanan(RSDetail!Jumlah, "###,###,### +")
    Else
        'jika satu digit awalnya kode perkiraannya 1 maka deri tanda minus (-)
        Tampilkan.Print Tab(25); RKanan(RSDetail!Jumlah, "###,###,### -")
    End If
    RSDetail.MoveNext
Loop
'End If
Tampilkan.Print Tab(5); MGrs
Tampilkan.Print Tab(5); "Pendapatan :";
'cetak pendatapan
Tampilkan.Print Tab(25); RKanan(RSGaji!Pendapatan, "###,###,### +");
Tampilkan.Print Tab(5); "Potongan   :";
'cetak potongan
Tampilkan.Print Tab(25); RKanan(RSGaji!Potongan, "###,###,### -");
Tampilkan.Print Tab(5); MGrs
Tampilkan.Print Tab(5); "Total      :";
'cetak totalnya (pendapatan -potongan)
If RSGaji!Pendapatan = RSGaji!Potongan Then
    Tampilkan.Print Tab(34); RSGaji!Pendapatan - RSGaji!Potongan
Else
    Tampilkan.Print Tab(25); RKanan(RSGaji!Pendapatan - RSGaji!Potongan, "###,###,### +");
End If
Tampilkan.Print Tab(5); MGrs
Tampilkan.Print
Tampilkan.Print Tab(5); "ESC    = Tutup Form Struk Gaji"
Tampilkan.Print Tab(5); "Enter  = Cetak Ke Printer"
End Sub

'meratakan angka di kanan
Private Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

