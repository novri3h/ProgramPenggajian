VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama Program Penggajian"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5550
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
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   3975
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2040
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1A7F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1AB12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1AE2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1B146
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1B460
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1B77A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1BA94
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1BDAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1C0C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "File"
      Begin VB.Menu mnkasir 
         Caption         =   "Kasir"
      End
      Begin VB.Menu mnperkiraan 
         Caption         =   "Perkiraan"
      End
      Begin VB.Menu mnpegawai 
         Caption         =   "Pegawai"
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnpenggajian 
         Caption         =   "Penggajian"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlappegawai 
         Caption         =   "Data Pegawai"
      End
      Begin VB.Menu mnlapperkiraan 
         Caption         =   "Data Perkiraan"
      End
      Begin VB.Menu mnlapgaji 
         Caption         =   "Penggajian"
      End
      Begin VB.Menu mnrincian 
         Caption         =   "Rincian Penggajian"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
'jika menekan ESC, munculkan pesan
If KeyAscii = 27 Then End
If KeyAscii = 13 Then Penggajian.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnkasir_Click()
Kasir.Show
End Sub

Private Sub mnkeluar_Click()
End
End Sub

Private Sub mnlapgaji_Click()
Laporan.Show
End Sub

Private Sub mnlappegawai_Click()
'panggil file laporan
CrystalReport1.ReportFileName = App.Path & "\Lap Pegawai.rpt"
'jika ada perubahan data direfresh
CrystalReport1.WindowState = crptMaximized
'tampilkan satu layar penuh
CrystalReport1.RetrieveDataFiles
'tampilkan ke layar
CrystalReport1.Action = 0
End Sub

Private Sub mnlapperkiraan_Click()
CrystalReport1.ReportFileName = App.Path & "\Lap Perkiraan.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 0
End Sub

Private Sub mnpegawai_Click()
Pegawai.Show
End Sub

Private Sub mnpenggajian_Click()
Penggajian.Show
End Sub

Private Sub mnperkiraan_Click()
Perkiraan.Show
End Sub

Private Sub mnsql_Click()
UjiCobaSQL.Show
End Sub

Private Sub mnrincian_Click()
Rincian.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "a"
        Kasir.Show
    Case "b"
        Perkiraan.Show
    Case "c"
        Pegawai.Show
    Case "d"
        Penggajian.Show
    Case "e"
        'panggil file laporan
CrystalReport1.ReportFileName = App.Path & "\Lap Pegawai.rpt"
'jika ada perubahan data direfresh
CrystalReport1.WindowState = crptMaximized
'tampilkan satu layar penuh
CrystalReport1.RetrieveDataFiles
'tampilkan ke layar
CrystalReport1.Action = 0

    Case "f"
       CrystalReport1.ReportFileName = App.Path & "\Lap Perkiraan.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 0
    Case "g"
        Laporan.Show
    Case "h"
        Rincian.Show
    Case "i"
        
        End
End Select
End Sub
