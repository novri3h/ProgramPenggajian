VERSION 5.00
Begin VB.Form Tampilkan 
   BackColor       =   &H80000009&
   Caption         =   "Struk Gaji Karyawan"
   ClientHeight    =   5520
   ClientLeft      =   -45
   ClientTop       =   240
   ClientWidth     =   4410
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
   ScaleHeight     =   5520
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Tampilkan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'program ini hampir sama dengan program transaksi penggajian
'bedanya adalah pencetakan dilakukan ke printer
'cetak ke printer cukup dengan menekan tombol ENTER

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Pesan = MsgBox("Printer sudah siap", vbYesNo)
    If Pesan = vbYes Then Cetakprinter
Else
    Unload Me
End If
End Sub


Sub Cetakprinter()
Dim MGrs As String
Printer.Font = "Courier New"
'cari data gaji dengan nomor slip terakhir
Call BukaDB
RSGaji.Open "select * from Gaji Where NomorSlp In(Select Max(NomorSlp)From Gaji)Order By NomorSlp Desc", Conn
'cari data kasir yang kodenya ada di tabel gaji
RSKasir.Open "select * from kasir where kodeksr='" & RSGaji!KodeKsr & "'", Conn
RSKasir.Requery
'cari data pegawai yang nip-nya ada di tabel gaji
RSPegawai.Open "select * from pegawai where NIP='" & RSGaji!NIP & "'", Conn
RSPegawai.Requery
Printer.Print
Printer.FontBold = True
Printer.Print
Printer.FontBold = False
'cetak nomor slip,tanggal,nip dan seterusnya
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print Tab(5); "Nomor Slip :  "; RSGaji!NomorSlp
Printer.Print Tab(5); "Tanggal    :  "; Format(RSGaji!Tanggal, "DD-MMMM-YYYY")
Printer.Print Tab(5); "NIP        :  "; RSGaji!NIP
Printer.Print Tab(5); "Nama       :  "; RSPegawai!NamaPgw
Printer.Print Tab(5); "Kasir      :  "; RSKasir!NamaKsr
MGrs = String$(33, "-")
'cetak isi tabel detailgaji berdasarkan nomor slip terakhir
Printer.Print Tab(5); MGrs
RSDetail.Open "select * from DetailGaji Where NomorSlp In(Select max(NomorSlp)From DetailGaji)", Conn
RSDetail.Requery
'If Not RSDetail.EOF Then
RSDetail.MoveFirst
Do While Not RSDetail.EOF
    'Call BukaDB
    Set RSPerkiraan = New ADODB.Recordset
    RSPerkiraan.Open "select * from Perkiraan where KodePrk='" & RSDetail!KodePrk & "'", Conn
    Printer.Print Tab(5); RSPerkiraan!NamaPrk;
    If Left(RSDetail!KodePrk, 1) = "0" Then
        'jika satu digit awal kode perkirannya 0 maka beri tanda plus (+)
        Printer.Print Tab(25); RKanan(RSDetail!Jumlah, "###,###,### +")
    Else
        'jika satu digit awalnya kode perkiraannya 1 maka deri tanda minus (-)
        Printer.Print Tab(25); RKanan(RSDetail!Jumlah, "###,###,### -")
    End If
    RSDetail.MoveNext
Loop
'End If
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Pendapatan :";
'cetak pendatapan
Printer.Print Tab(25); RKanan(RSGaji!Pendapatan, "###,###,### +");
Printer.Print Tab(5); "Potongan   :";
'cetak potongan
Printer.Print Tab(25); RKanan(RSGaji!Potongan, "###,###,### -");
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Total      :";
'cetak totalnya (pendapatan -potongan)
If RSGaji!Pendapatan = RSGaji!Potongan Then
    Printer.Print Tab(34); RSGaji!Pendapatan - RSGaji!Potongan
Else
    Printer.Print Tab(25); RKanan(RSGaji!Pendapatan - RSGaji!Potongan, "###,###,### +");
End If
Printer.Print Tab(5); MGrs
Printer.Print
Printer.EndDoc
End Sub

'meratakan angka di kanan
Private Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

