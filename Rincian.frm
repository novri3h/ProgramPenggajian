VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Rincian 
   Caption         =   "Rincian dan Cetak Gaji"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4710
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
   ScaleHeight     =   5055
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCetak 
      Caption         =   "Cetak Gaji"
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox TxtNIP 
      Enabled         =   0   'False
      Height          =   320
      Left            =   1080
      TabIndex        =   14
      Top             =   480
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox TxtBagian 
      Enabled         =   0   'False
      Height          =   320
      Left            =   1080
      TabIndex        =   6
      Top             =   1200
      Width           =   3500
   End
   Begin VB.TextBox TxtNama 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   3500
   End
   Begin VB.TextBox TxtPendapatan 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   320
      Left            =   3000
      TabIndex        =   4
      Top             =   3960
      Width           =   1250
   End
   Begin VB.TextBox TxtPotongan 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   320
      Left            =   3000
      TabIndex        =   3
      Top             =   4320
      Width           =   1250
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   320
      Left            =   3000
      TabIndex        =   2
      Top             =   4680
      Width           =   1250
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Rincian.frx":0000
      Height          =   2300
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   4048
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
         DataField       =   "Nama Pembayaran"
         Caption         =   "Nama Pembayaran"
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
            ColumnWidth     =   2805,166
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DTDetail 
      Height          =   345
      Left            =   120
      Top             =   3960
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "Detail"
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
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bagian"
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NIP"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Slip"
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
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   900
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
      Left            =   1920
      TabIndex        =   10
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Potongan"
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   4320
      Width           =   1005
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Gaji Bersih"
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   4680
      Width           =   1005
   End
End
Attribute VB_Name = "Rincian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'program ini hampir sama dgn pencetakan slip gaji
'bedanya terletak pada pilihan nomor slip dan nip

Private Sub CmdCetak_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Combo1_Keypress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Call BukaDB
Combo1.Clear
RSGaji.Open "Select Distinct NomorSlp from Gaji", Conn
Do Until RSGaji.EOF
    Combo1.AddItem RSGaji!NomorSlp
    RSGaji.MoveNext
Loop
Conn.Close
End Sub

Private Sub Combo1_Click()
Call BukaDB
RSGaji.Open "select * from Gaji where NomorSlp='" & Combo1.Text & "'", Conn
RSPegawai.Open "select * from Pegawai where NIP='" & RSGaji!NIP & "'", Conn
If Not RSPegawai.EOF Then
    TxtNIP = RSPegawai!NIP
    TxtNama = RSPegawai!NamaPgw
    TxtBagian = RSPegawai!Bagian
End If

DTDetail.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOGaji.mdb"
DTDetail.RecordSource = "select NamaPrk as [Nama Pembayaran],Jumlah from Perkiraan,DetailGaji where DetailGaji.kodePrk=Perkiraan.kodePrk and NomorSlp='" & Combo1.Text & "'"
Set DataGrid1.DataSource = DTDetail
DTDetail.Refresh
DataGrid1.Refresh

TxtPendapatan = Format(GajiKotor, "##,###,###")
TxtPotongan = Format(Potongan, "##,###,###")
TxtTotal = Format(GajiKotor - Potongan, "##,###,###")
Conn.Close
End Sub

Function GajiKotor()
    Set RSGajiKotor = New ADODB.Recordset
    RSGajiKotor.Open "Select sum(Jumlah) as TTLPendapatan from DetailGaji where left(kodeprk,1)=0 and NomorSlp='" & Combo1 & "'", Conn
    GajiKotor = RSGajiKotor!TTLPendapatan
End Function

Function Potongan()
    Set RSPotongan = New ADODB.Recordset
    RSPotongan.Open "Select sum(Jumlah) as TTLPotongan from DetailGaji where left(kodeprk,1)=1 and NomorSlp='" & Combo1 & "'", Conn
    Potongan = RSPotongan!TTLPotongan
End Function

Private Sub Bersihkan()
TxtNamaPgw = ""
TxtBagian = ""
TxtPendapatan = ""
TxtPotongan = ""
TxtTotal = ""
End Sub

Private Sub CmdBatal_Click()
Bersihkan
'Form_Activate
End Sub

Private Sub CmdTutup_Click()
CmdBatal_Click
Unload Me
End Sub

Private Sub CmdCetak_click()
If Combo1 = "" Then
    MsgBox "Nomor slip gaji kosong"
    Exit Sub
    Combo1.SetFocus
End If
Pesan = MsgBox("Printer sudah siap..?", vbYesNo, "Konfirmasi")
If Pesan = vbYes Then

Dim MGrs As String
Printer.Font = "Courier New"

Call BukaDB
RSGaji.Open "select * from Gaji Where NomorSlp ='" & Combo1 & "'", Conn
RSKasir.Open "select * from kasir where kodeksr='" & RSGaji!KodeKsr & "'", Conn
RSKasir.Requery
RSPegawai.Open "select * from pegawai where NIP='" & RSGaji!NIP & "'", Conn
RSPegawai.Requery
Printer.Print
Printer.FontBold = True
Printer.Print
Printer.FontBold = False

Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print Tab(5); "Nomor Slip :  "; RSGaji!NomorSlp
Printer.Print Tab(5); "Tanggal    :  "; Format(RSGaji!Tanggal, "DD-MMMM-YYYY")
Printer.Print Tab(5); "NIP        :  "; RSGaji!NIP
Printer.Print Tab(5); "Nama       :  "; RSPegawai!NamaPgw
Printer.Print Tab(5); "Kasir      :  "; RSKasir!NamaKsr
MGrs = String$(33, "-")
Printer.Print Tab(5); MGrs

RSDetail.Open "select * from DetailGaji Where NomorSlp ='" & Combo1 & "'", Conn
RSDetail.Requery
RSDetail.MoveFirst
Do While Not RSDetail.EOF
    Set RSPerkiraan = New ADODB.Recordset
    RSPerkiraan.Open "select * from Perkiraan where KodePrk='" & RSDetail!KodePrk & "'", Conn
    Printer.Print Tab(5); RSPerkiraan!NamaPrk;
    If Left(RSDetail!KodePrk, 1) = "0" Then
        Printer.Print Tab(25); RKanan(RSDetail!Jumlah, "###,###,### +")
    Else
        Printer.Print Tab(25); RKanan(RSDetail!Jumlah, "###,###,### -")
    End If
    RSDetail.MoveNext
Loop
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Pendapatan :";
Printer.Print Tab(25); RKanan(RSGaji!Pendapatan, "###,###,### +");
Printer.Print Tab(5); "Potongan   :";
Printer.Print Tab(25); RKanan(RSGaji!Potongan, "###,###,### -");
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Total      :";
If RSGaji!Pendapatan = RSGaji!Potongan Then
    Printer.Print Tab(34); RSGaji!Pendapatan - RSGaji!Potongan
Else
    Printer.Print Tab(25); RKanan(RSGaji!Pendapatan - RSGaji!Potongan, "###,###,### +");
End If
Printer.Print Tab(5); MGrs
Printer.Print
Printer.EndDoc
End If
End Sub

Private Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function
