VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Sistem Informasi Pengelolaan Pajak Bumi dan Bangunan Perdesaan dan Perkotaan (PBB-P2)"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   13755
   Icon            =   "frmUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmUtama.frx":1CCA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6270
      Top             =   3840
   End
   Begin MSComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7785
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   7064
            Text            =   "Selamat Datang di Aplikasi SIMDA-PBB"
            TextSave        =   "Selamat Datang di Aplikasi SIMDA-PBB"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   3175
            MinWidth        =   3175
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog gLOK 
      Left            =   2475
      Top             =   3705
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu mnSearch 
         Caption         =   "Search"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnCT 
         Caption         =   "&Duplikat Data"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnPulih 
         Caption         =   "&Pemulihan Data"
         Shortcut        =   ^Z
      End
      Begin VB.Menu G0 
         Caption         =   "-"
      End
      Begin VB.Menu mnBackop 
         Caption         =   "Pengembalian Objek Pajak"
         Begin VB.Menu mnAwal 
            Caption         =   "Data KPP P&ratama Tahun 2013"
         End
         Begin VB.Menu mnExist 
            Caption         =   "Data E&xist"
         End
      End
      Begin VB.Menu OP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnLog 
         Caption         =   "Log &Off"
         Shortcut        =   ^Q
      End
      Begin VB.Menu G2 
         Caption         =   "-"
      End
      Begin VB.Menu mnKeluar 
         Caption         =   "K&eluar"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnDaftar 
      Caption         =   "Penda&taan"
      Begin VB.Menu mnZNT 
         Caption         =   "&Zona Nilai Tanah (ZNT)"
         Begin VB.Menu mnBlok 
            Caption         =   "Blok"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnZNT1 
            Caption         =   "ZNT"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnNIR 
            Caption         =   "NIR"
            Shortcut        =   ^I
         End
         Begin VB.Menu G9 
            Caption         =   "-"
         End
         Begin VB.Menu nmJalan 
            Caption         =   "&Nama Jalan"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu mnDBKB 
         Caption         =   "&DBKB"
         Begin VB.Menu mndStandard 
            Caption         =   "DBKB S&tandar"
            Begin VB.Menu mnDBKBOtomatis 
               Caption         =   "DBKB Utama dan Material (*Sistem)"
               Shortcut        =   ^{F4}
            End
            Begin VB.Menu c2 
               Caption         =   "-"
            End
            Begin VB.Menu mnDBKB3 
               Caption         =   "Custom: DBKB Fasilitas (*Wajib)"
               Shortcut        =   ^{F5}
            End
            Begin VB.Menu c1 
               Caption         =   "-"
            End
            Begin VB.Menu mnDBKD 
               Caption         =   "Custom: DBKB Utama (Option)"
               Shortcut        =   ^{F6}
            End
            Begin VB.Menu mnDBKB1 
               Caption         =   "Custom: DBKB Material (Option)"
               Shortcut        =   ^{F9}
            End
         End
         Begin VB.Menu G1a 
            Caption         =   "-"
         End
         Begin VB.Menu mnDBKB2 
            Caption         =   "DBKB Non Standar"
            Shortcut        =   ^{F7}
         End
         Begin VB.Menu mnSistem3_8 
            Caption         =   "DBKB JPB3_JPB8 (*Sistem)"
            Shortcut        =   %{BKSP}
         End
      End
      Begin VB.Menu G1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSubjek 
         Caption         =   "&Subjek Pajak"
         Shortcut        =   ^J
      End
      Begin VB.Menu G7 
         Caption         =   "-"
      End
      Begin VB.Menu mnOPajak 
         Caption         =   "&Objek Pajak"
         Begin VB.Menu mnOPBumi 
            Caption         =   "Bu&mi"
            Shortcut        =   ^T
         End
         Begin VB.Menu G8 
            Caption         =   "-"
         End
         Begin VB.Menu mnOPBangunan 
            Caption         =   "Ban&gunan"
            Shortcut        =   ^B
         End
      End
   End
   Begin VB.Menu mnNilai 
      Caption         =   "&Penilaian"
      Begin VB.Menu mnNilai1 
         Caption         =   "Penilaian Ma&ssal"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnLaporanN 
         Caption         =   "Laporan Penilaian"
         Begin VB.Menu mnBefore 
            Caption         =   "&Perbandingan Dengan Tahun Sebelumnya"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnG21 
            Caption         =   "-"
         End
         Begin VB.Menu mnLIndividu 
            Caption         =   "Bangunan Secara &Individu"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mngB1 
            Caption         =   "-"
         End
         Begin VB.Menu mnLMassal1 
            Caption         =   "&Bumi Secara Massal"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnLMassal2 
            Caption         =   "&Bangunan Secara Massal"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnG2 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnLBB1 
            Caption         =   "Laporan Penilaian Bumi dan Bangunan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnBB2 
            Caption         =   "Laporan Penilaian Tahun Sebelumnya"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnPenetapan 
      Caption         =   "Pe&netapan"
      Begin VB.Menu mnMinimal 
         Caption         =   "&PBB Minimal"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu GG9 
         Caption         =   "-"
      End
      Begin VB.Menu mnNJOPTKP 
         Caption         =   "Penetapan NJOPT&KP"
         Shortcut        =   ^K
      End
      Begin VB.Menu gNJOPTKP 
         Caption         =   "-"
      End
      Begin VB.Menu mnPBB1 
         Caption         =   "Penetapan SPPT &Massal"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnPBB2 
         Caption         =   "Penetapan SPPT &Tunggal"
         Shortcut        =   ^L
      End
      Begin VB.Menu G6 
         Caption         =   "-"
      End
      Begin VB.Menu mnLunas1 
         Caption         =   "Pelunasa&n Massal"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnLunas2 
         Caption         =   "Pelunasan &Tunggal"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnReferensi 
      Caption         =   "&Referensi"
      Begin VB.Menu mnWilayah 
         Caption         =   "&Wilayah"
         Begin VB.Menu mnKec 
            Caption         =   "Ke&camatan"
            Shortcut        =   ^C
         End
         Begin VB.Menu mnKel 
            Caption         =   "Ke&lurahan/Desa"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnMekar 
            Caption         =   "Pemekaran"
            Shortcut        =   ^{F1}
            Visible         =   0   'False
         End
      End
      Begin VB.Menu G4 
         Caption         =   "-"
      End
      Begin VB.Menu mnKlas1 
         Caption         =   "Klasifikasi Tarif/Tanah/Bangunan"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu G3 
         Caption         =   "-"
      End
      Begin VB.Menu mnBayar 
         Caption         =   "Tempat Pemba&yaran"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu y1 
         Caption         =   "-"
      End
      Begin VB.Menu mnUlin 
         Caption         =   "Status Penggunaan Ka&yu Ulin"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu gl1 
         Caption         =   "-"
      End
      Begin VB.Menu mnResource 
         Caption         =   "Update Resources"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnTarif 
         Caption         =   "Tari&f"
         Visible         =   0   'False
      End
      Begin VB.Menu mnBangunan 
         Caption         =   "Kelas Bang&unan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnBumi 
         Caption         =   "Kelas Bu&mi"
         Visible         =   0   'False
      End
      Begin VB.Menu gPosting 
         Caption         =   "-"
      End
      Begin VB.Menu mnPosting 
         Caption         =   "Posting SPPT Lama"
      End
   End
   Begin VB.Menu mnLaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnwSPPT 
         Caption         =   "SPP&T"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnwSSPD 
         Caption         =   "&SSPD"
         Shortcut        =   ^R
      End
      Begin VB.Menu G19 
         Caption         =   "-"
      End
      Begin VB.Menu mnDHKP 
         Caption         =   "&DHKP"
         Shortcut        =   ^V
      End
      Begin VB.Menu G20 
         Caption         =   "-"
      End
      Begin VB.Menu mnNew 
         Caption         =   "&Format Lama"
         Begin VB.Menu mnSPPT 
            Caption         =   "Sura&t Pemberitahuan Pajak Terutang (SPPT)"
            Shortcut        =   +{F11}
         End
         Begin VB.Menu mnSTTS 
            Caption         =   "Surat Seto&ran Pajak Daerah (SSPD)"
            Shortcut        =   +{F12}
         End
         Begin VB.Menu mnLDHKP 
            Caption         =   "&Daftar Himpunan Ketetapan Pajak (DHKP)"
            Shortcut        =   +^{F1}
         End
         Begin VB.Menu G21 
            Caption         =   "-"
         End
         Begin VB.Menu mntunggal1 
            Caption         =   "SPP&T Tunggal"
         End
         Begin VB.Menu mntunggal2 
            Caption         =   "SSP&D Tunggal"
         End
      End
      Begin VB.Menu MNg5 
         Caption         =   "-"
      End
      Begin VB.Menu mnLap1 
         Caption         =   "Klasifi&kasi dan Besaran NJOP Bumi"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnLDBKB 
         Caption         =   "Da&ftar Komponen Biaya Bangunan (DBKB)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MNg4 
         Caption         =   "-"
      End
      Begin VB.Menu MNSIMULASI 
         Caption         =   "&Simulasi Laporan Penetapan SPPT"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnUtility 
      Caption         =   "&Utility"
      Begin VB.Menu mnPejabat 
         Caption         =   "Pe&jabat"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mj 
         Caption         =   "-"
      End
      Begin VB.Menu mnWewenang 
         Caption         =   "Wewenang"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnUser 
         Caption         =   "&User"
         Shortcut        =   ^Y
      End
      Begin VB.Menu MNg1 
         Caption         =   "-"
      End
      Begin VB.Menu mnCopyDB 
         Caption         =   "Backup &Database"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnGSP 
         Caption         =   "-"
      End
      Begin VB.Menu mnSetting 
         Caption         =   "Setting Printe&r"
         Shortcut        =   ^U
      End
      Begin VB.Menu gs1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSet 
         Caption         =   "&Setting Laporan"
      End
      Begin VB.Menu gLog 
         Caption         =   "-"
      End
      Begin VB.Menu mnHLog 
         Caption         =   "&Log Data"
         Begin VB.Menu mnPrint 
            Caption         =   "&Cetak Log Data Objek"
         End
         Begin VB.Menu mnLogBaru 
            Caption         =   "C&etak Log Ketetapan Baru"
         End
         Begin VB.Menu GHA 
            Caption         =   "-"
         End
         Begin VB.Menu mnLBackup 
            Caption         =   "&Backup Log"
            Visible         =   0   'False
         End
         Begin VB.Menu mnHapusLog 
            Caption         =   "&Hapus Log Data"
         End
      End
   End
   Begin VB.Menu mnOFF1 
      Caption         =   "Lo&g off"
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Activate()
On Error Resume Next
mnSetting.Enabled = False
stsBar.Panels.Item(4).Text = Format(Now, "dddd, DD-MM-YYYY")

xLoad = 1

cekTampil = 1
mnSet.Checked = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
TANYA = MsgBox("Apa anda yakin ingin keluar dari Aplikasi?", vbQuestion + vbYesNo, "Keluar")
             If TANYA = vbYes Then
                Kill "A:\*.TMP"
                Kill "B:\*.TMP"
                Kill "D:\*.TMP"
                Kill "C:\*.TMP"
                Kill "E:\*.TMP"
                Kill "F:\*.TMP"
                Kill "G:\*.TMP"
                Kill "H:\*.TMP"
                Kill "I:\*.TMP"
                  Kill App.Path & "\*.TMP"
                End
            Else
                Cancel = True
                Exit Sub
            End If
End Sub

Private Sub mnAdmin_Click()
On Error Resume Next
frmPajak.Show
End Sub

Private Sub mnAwal_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
TANYA = MsgBox("Aplikasi akan mengembalikan data objek pajak" & _
            vbCrLf & "Tahun 2013 (KPP-Pratama), [Data Terbaru terkait Penghapusan/" & _
            vbCrLf & "Pemutakhiran/Penambahan] dianggap tidak ada, Lanjut?", vbQuestion + vbYesNo, "Restored...!")
If TANYA = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
Q_STR = "BACK_OP '1'"
openDB (Q_STR)
MsgBox "Data Objek Pajak Tahun 2013 berhasil di restore..!"
Screen.MousePointer = vbDefault

End Sub

Private Sub mnBangunan_Click()
On Error Resume Next
frmKelas_Bangunan.Show
End Sub

Private Sub mnBayar_Click()
On Error Resume Next
frmBank.Show
End Sub

Private Sub mnBB2_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 50

End Sub

Private Sub mnBefore_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 60

End Sub

Private Sub mnBlok_Click()
On Error Resume Next
frmBlok.Show
End Sub

Private Sub mnBumi_Click()
On Error Resume Next
frmKelas_Bumi.Show
End Sub

Private Sub mnCopyDB_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
masuk = MsgBox("Anda harus menentukan lokasi penyimpanan" & _
            vbCrLf & "file backup. LANJUT?", vbYesNo + vbInformation, "Backup...!")
If masuk = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
Screen.MousePointer = vbDefault
With gLOK
        .DialogTitle = "Pilih Lokasi Database dan Ketik Nama File Backup"
        .CancelError = False
        '.Filter = "All Files (*.*)|*.*"
        '.Filter = "File Access(*.MDB)|*.MDB"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        If UCase(Right(Trim(.FileName), 4)) = ".BAK" Then
            xxTahun = .FileName
        Else
            xxTahun = .FileName & ".Bak"
        End If
        
End With
 Screen.MousePointer = vbHourglass
CTEMU = GetAttr(xxTahun)
If CTEMU = 32 Then
    MsgBox "Backup File " & xxTahun & " Gagal!" & _
            vbCrLf & "Silahkan Ganti Nama Lain.", vbCritical, "Fail"
    GoTo Keluar
End If

TANYA = MsgBox("Apa anda yakin membuat backup database, " & _
                vbCrLf & "dengan nama file : " & xxTahun, vbQuestion + vbYesNo, "Backup...")
If TANYA = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
c_SQL = " backup database dbpajak to disk = '" + xxTahun + "'  with format,   Medianame='MN_PBB',  Name='Full Backup of dbPajak'"
openDB (c_SQL)
MsgBox "Backup Database Sukses...!", vbInformation, "Sukses!"
Keluar:
Screen.MousePointer = vbDefault
End Sub

Private Sub mnCT_Click()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
'MENDUPLIKASI ISI TABEL
TANYA = MsgBox("Aplikasi akan menduplikat struktur dan isi beberapa tabel" & _
            vbCrLf & "Lanjut?", vbQuestion + vbYesNo, "Copied...!")
If TANYA = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
''Duplikasi Tabel DAT_OP_BUMI
'C_TAB = "SELECT * INTO C_BUMI FROM DAT_OP_BUMI"
'openDB (C_TAB)
''Duplikasi Tabel DAT_OP_BANGUNAN
'C_TAB = "SELECT * INTO C_BANGUNAN FROM DAT_OP_BANGUNAN"
'openDB (C_TAB)
''Duplikasi Tabel DAT_OBJEK_PAJAK
'C_TAB = "SELECT * INTO C_OBJEK FROM DAT_OBJEK_PAJAK"
'openDB (C_TAB)
''Duplikasi Tabel DAT_NILAI_INDIVIDU
'C_TAB = "SELECT * INTO C_INDIVIDU FROM DAT_NILAI_INDIVIDU"
'openDB (C_TAB)
'
''Duplikasi Tabel DAT_ZNT
'C_TAB = "SELECT * INTO C_ZNT FROM DAT_ZNT"
'openDB (C_TAB)
'
''Duplikasi Tabel DAT_FASILITAS_BANGUNAN
'C_TAB = "SELECT * INTO C_FASILITAS FROM DAT_FASILITAS_BANGUNAN"
'openDB (C_TAB)
'
''Duplikasi Tabel DAT_SUBJEK_PAJAK_NJOPTKP
'C_TAB = "SELECT * INTO C_NJOPTKP FROM DAT_SUBJEK_PAJAK_NJOPTKP"
'openDB (C_TAB)
'
''Duplikasi JPB
'
''Duplikasi Tabel JPB2
'C_TAB = "SELECT * INTO C2 FROM DAT_JPB2"
'openDB (C_TAB)
''Duplikasi Tabel JPB3
'C_TAB = "SELECT * INTO C3 FROM DAT_JPB3"
'openDB (C_TAB)
''Duplikasi Tabel JPB4
'C_TAB = "SELECT * INTO C4 FROM DAT_JPB4"
'openDB (C_TAB)
''Duplikasi Tabel JPB5
'C_TAB = "SELECT * INTO C5 FROM DAT_JPB5"
'openDB (C_TAB)
''Duplikasi Tabel JPB6
'C_TAB = "SELECT * INTO C6 FROM DAT_JPB6"
'openDB (C_TAB)
''Duplikasi Tabel JPB7
'C_TAB = "SELECT * INTO C7 FROM DAT_JPB7"
'openDB (C_TAB)
''Duplikasi Tabel JPB8
'C_TAB = "SELECT * INTO C8 FROM DAT_JPB8"
'openDB (C_TAB)
''Duplikasi Tabel JPB9
'C_TAB = "SELECT * INTO C9 FROM DAT_JPB9"
'openDB (C_TAB)
''Duplikasi Tabel JPB12
'C_TAB = "SELECT * INTO C12 FROM DAT_JPB12"
'openDB (C_TAB)
''Duplikasi Tabel JPB13
'C_TAB = "SELECT * INTO C13 FROM DAT_JPB13"
'openDB (C_TAB)
''Duplikasi Tabel JPB14
'C_TAB = "SELECT * INTO C14 FROM DAT_JPB14"
'openDB (C_TAB)
''Duplikasi Tabel JPB15
'C_TAB = "SELECT * INTO C15 FROM DAT_JPB15"
'openDB (C_TAB)
''Duplikasi Tabel JPB16
'C_TAB = "SELECT * INTO C16 FROM DAT_JPB16"
'openDB (C_TAB)
''Duplikasi Tabel JPB17
'C_TAB = "SELECT * INTO C17 FROM DAT_JPB17"
'openDB (C_TAB)
C_STR = "BUAT_DUPLIKASI"
openDB (C_STR)

MsgBox "Proses Duplikasi Berhasil...!", vbInformation, "SUKSES!"
Screen.MousePointer = vbDefault
Salah:
If Err.Number = 0 Then
    Exit Sub
ElseIf Err.Number = -2147217900 Then
'    tt1 = MsgBox("Tabel sudah berisi, apa ingin ditimpa?", vbQuestion + vbYesNo, "Exist...!")
'    If tt1 = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
'    'Menghapus data yang sudah ada pada tabel DAT_OP_BUMI
'    C_TAB = "DELETE  from C_BUMI"
'    openDB (C_TAB)
'    'Mengembalikan nilai Sebelumnya
'    C_TAB = "INSERT INTO C_BUMI SELECT * FROM DAT_OP_BUMI "
'    openDB (C_TAB)
'    'Menghapus Tabel yang ada pada tabel DAT_OP_BANGUNAN
'    C_TAB = "DELETE  FROM C_BANGUNAN"
'    openDB (C_TAB)
'    'Mengembalikan nilai Sebelumnya
'    C_TAB = "INSERT INTO C_BANGUNAN SELECT * FROM DAT_OP_BANGUNAN"
'    openDB (C_TAB)
'    'Menghapus data yang sudah ada pada Tabel DAT_OBJEK_PAJAK
'    C_TAB = "DELETE  FROM C_OBJEK "
'    openDB (C_TAB)
'    'Mengembalikan nilai Sebelumnya
'    C_TAB = "INSERT INTO C_OBJEK SELECT * FROM DAT_OBJEK_PAJAK"
'    openDB (C_TAB)
'    'Menghapus data yang sudah ada pada Tabel DAT_NILAI_INDIVIDU
'    C_TAB = "DELETE  FROM C_INDIVIDU"
'    openDB (C_TAB)
'    'Mengembalikan nilai Sebelumnya
'    C_TAB = "INSERT INTO C_INDIVIDU SELECT * FROM DAT_NILAI_INDIVIDU "
'    openDB (C_TAB)
'    'Duplikasi Tabel DAT_ZNT
'    C_TAB = "DELETE  FROM C_ZNT"
'    openDB (C_TAB)
'    'Mengembalikan nilai Sebelumnya
'    C_TAB = "INSERT INTO C_ZNT SELECT * FROM DAT_ZNT"
'    openDB (C_TAB)
'    'Duplikasi Tabel DAT_FASILITAS_BANGUNAN
'    C_TAB = "DELETE  FROM  C_FASILITAS "
'    openDB (C_TAB)
'    'Mengembalikan nilai Sebelumnya
'    C_TAB = "INSERT INTO C_FASILITAS SELECT * FROM DAT_FASILITAS_BANGUNAN"
'    openDB (C_TAB)
'    'Duplikasi Tabel DAT_SUBJEK_PAJAK_NJOPTKP
'    C_TAB = "DELETE  FROM C_NJOPTKP"
'    openDB (C_TAB)
'    'Mengembalikan nilai Sebelumnya
'    C_TAB = "INSERT INTO C_NJOPTKP SELECT * FROM DAT_SUBJEK_PAJAK_NJOPTKP"
'    openDB (C_TAB)
'        HAPUS_TAB
'        'Duplikasi Tabel JPB2
'        C_TAB = "INSERT INTO C2 select * FROM DAT_JPB2"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB3
'        C_TAB = "INSERT INTO C3 select * FROM DAT_JPB3"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB4
'        C_TAB = "INSERT INTO C4 select * FROM DAT_JPB4"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB5
'        C_TAB = "INSERT INTO C5 select * FROM DAT_JPB5"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB6
'        C_TAB = "INSERT INTO C6 select * FROM DAT_JPB6"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB7
'        C_TAB = "INSERT INTO C7 select * FROM DAT_JPB7"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB8
'        C_TAB = "INSERT INTO C8 select * FROM DAT_JPB8"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB9
'        C_TAB = "INSERT INTO C9 select * FROM DAT_JPB9"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB12
'        C_TAB = "INSERT INTO C12 select * FROM DAT_JPB12"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB13
'        C_TAB = "INSERT INTO C13 select * FROM DAT_JPB13"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB14
'        C_TAB = "INSERT INTO C14 select * FROM DAT_JPB14"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB15
'        C_TAB = "INSERT INTO C15 select * FROM DAT_JPB15"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB16
'        C_TAB = "INSERT INTO C16 select * FROM DAT_JPB16"
'        openDB (C_TAB)
'        'Duplikasi Tabel JPB17
'        C_TAB = "INSERT INTO C17 select * FROM DAT_JPB17"
'        openDB (C_TAB)
    C_STR = "HAPUS_DUPLIKASI"
    openDB (C_STR)
    MsgBox "Proses Duplikasi Berhasil...!", vbInformation, "SUKSES!"
Else
    MsgBox Err.Number & " : " & Err.Description, vbCritical, "Error"
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub mnDBKB1_Click()
On Error Resume Next
frmKomponenBiaya2.Show 'frmDBKB.Show
End Sub


Private Sub mnDBKB2_Click()
On Error Resume Next
frmJPB.Show
End Sub

Private Sub mnDBKB3_Click()
On Error Resume Next
frmKomponenBiaya1.Show
End Sub

Private Sub mnDBKBOtomatis_Click()
On Error Resume Next
frmBentuk_DBKB.Show
cBentuk = 1
End Sub

Private Sub mnDBKD_Click()
On Error Resume Next
frmKomponenBiaya3.Show
End Sub

Private Sub mnDHKP_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 300


End Sub

Private Sub mnExist_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
TANYA = MsgBox("Aplikasi akan mengembalikan data objek pajak" & _
            vbCrLf & "Tahun 2013, dikecualikan data baru dan penghapusan...?", vbQuestion + vbYesNo, "Restored...!")
If TANYA = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
Q_STR = "BACK_OP '2'"
openDB (Q_STR)
MsgBox "Data Objek Pajak berhasil di restore..!"
Screen.MousePointer = vbDefault
End Sub

Private Sub mnHapusLog_Click()
On Error Resume Next
Unload frmCetakLog
Unload rptPBB
'rptPBB.Show
frmCetakLog.Show
J_CETAK = 113
End Sub



Private Sub mnKec_Click()
On Error Resume Next
frmKec.Show
End Sub

Private Sub mnKel_Click()
On Error Resume Next
frmKel.Show
End Sub

Private Sub mnKeluar_Click()
On Error Resume Next
TANYA = MsgBox("Apa anda yakin ingin keluar dari Aplikasi?", vbQuestion + vbYesNo, "Keluar")
             If TANYA = vbYes Then
                Kill "A:\*.TMP"
                Kill "B:\*.TMP"
                Kill "D:\*.TMP"
                Kill "C:\*.TMP"
                Kill "E:\*.TMP"
                Kill "F:\*.TMP"
                Kill "G:\*.TMP"
                Kill "H:\*.TMP"
                Kill "I:\*.TMP"
                Kill App.Path & "\*.TMP"
                End
            Else
                Cancel = True
                Exit Sub
            End If
End Sub

Private Sub mnKlas1_Click()
On Error Resume Next
frmKlasifikasi.Show
End Sub

Private Sub mnLap1_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 4
End Sub

Private Sub mnLBB1_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 40

End Sub

Private Sub mnLDBKB_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 6
End Sub

Private Sub mnLDHKP_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 3
End Sub

Private Sub mnLIndividu_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 10
End Sub

Private Sub mnListSP_Click()
On Error Resume Next
frmList_Subjek.Show
End Sub

Private Sub mnLMassal2_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 30


End Sub
Private Sub mnLMassal1_Click()
On Error Resume Next

Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 20
End Sub

Private Sub mnLog_Click()
On Error Resume Next
xLoad = 0
frmLogin.Show
End Sub

Private Sub mnLogBaru_Click()
On Error Resume Next
Unload frmCetakLog
Unload rptPBB
'rptPBB.Show
frmCetakLog.Show
J_CETAK = 112
End Sub

Private Sub mnLunas1_Click()
On Error Resume Next
frmBayar2.Show
End Sub

Private Sub mnLunas2_Click()
On Error Resume Next
frmBayar1.Show
End Sub

Private Sub mnMinimal_Click()
On Error Resume Next
frmTarif.Show
End Sub

Private Sub mnNilai1_Click()
On Error Resume Next
frmNilai_Massal.Show
End Sub

Private Sub mnNIR_Click()
On Error Resume Next
frmNIR.Show
End Sub

Private Sub mnNJOPTKP_Click()
On Error Resume Next
ccMenu = 2
frmNJOPTKP.Show
End Sub

Private Sub mnOFF1_Click()
On Error Resume Next
frmLogin.Show
End Sub

Private Sub mnOPBangunan_Click()
On Error Resume Next
xID = ""
frmObjek_Pajak_Bg.Show
End Sub

Private Sub mnOPBumi_Click()
On Error Resume Next
xID = ""
frmObjek_Pajak_Bm.Show
End Sub

Private Sub mnPBB1_Click()
On Error Resume Next
frmSPPT_Massal.Show
End Sub

Private Sub mnPBB2_Click()
On Error Resume Next
frmSPPT_Tunggal.Show
End Sub

Private Sub mnPenetapanOP_Click()
On Error Resume Next
frmPBB.Show
End Sub

Private Sub mnPejabat_Click()
On Error Resume Next
frmPejabat.Show
End Sub

Private Sub mnPosting_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
'Memindahkan SPPT Lama
'ElseIf xxPro = "2" Then
'    bc_STR = "INSERT INTO SPPT_1 SELECT * FROM SPPT WHERE SPPT.KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' and SPPT.THN_PAJAK_SPPT='" & ccTahun.Text & "'"
'Else
'    bc_STR = "INSERT INTO SPPT_1 SELECT * FROM SPPT WHERE SPPT.KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' and SPPT.KD_KELURAHAN='" & Left(Trim(ccKel.Text), 3) & "' and SPPT.THN_PAJAK_SPPT='" & ccTahun.Text & "'"
'End If
    MsgBox "Posting SPPT Lama ini sebaiknya dibuat sebelum" & _
    vbCrLf & "pemutakhiran/penghapusan/penambahan SPOP", vbCritical + vbOKOnly, "Warning"

ccMenu = 1
frmNJOPTKP.Show
Screen.MousePointer = vbDefault
End Sub

Private Sub mnPrint_Click()
On Error Resume Next
Unload frmCetakLog
Unload rptPBB
'rptPBB.Show
frmCetakLog.Show
J_CETAK = 111
End Sub

Private Sub mnPulih_Click()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
TANYA = MsgBox("Sistem akan mengembalikan data lama " & _
     vbCrLf & "Seluruh data baru akan dihapus " & _
    vbCrLf & "Lanjut ?", vbQuestion + vbYesNo, "Reset..!")
If TANYA = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
'Menghapus data yang sudah ada pada tabel DAT_OP_BUMI
'C_TAB = "DELETE  from DAT_OP_BUMI"
'openDB (C_TAB)
''Mengembalikan nilai Sebelumnya
'C_TAB = "INSERT INTO DAT_OP_BUMI SELECT * FROM C_BUMI"
'openDB (C_TAB)
''Menghapus Tabel yang ada pada tabel DAT_OP_BANGUNAN
'C_TAB = "DELETE  FROM DAT_OP_BANGUNAN"
'openDB (C_TAB)
''Mengembalikan nilai Sebelumnya
'C_TAB = "INSERT INTO DAT_OP_BANGUNAN SELECT * FROM C_BANGUNAN"
'openDB (C_TAB)
''Menghapus data yang sudah ada pada Tabel DAT_OBJEK_PAJAK
'C_TAB = "DELETE  FROM DAT_OBJEK_PAJAK"
'openDB (C_TAB)
''Mengembalikan nilai Sebelumnya
'C_TAB = "INSERT INTO DAT_OBJEK_PAJAK SELECT * FROM C_OBJEK "
'openDB (C_TAB)
''Menghapus data yang sudah ada pada Tabel DAT_NILAI_INDIVIDU
'C_TAB = "DELETE  FROM DAT_NILAI_INDIVIDU"
'openDB (C_TAB)
''Mengembalikan nilai Sebelumnya
'C_TAB = "INSERT INTO DAT_NILAI_INDIVIDU SELECT * FROM C_INDIVIDU "
'openDB (C_TAB)
''Duplikasi Tabel DAT_ZNT
'C_TAB = "DELETE  FROM DAT_ZNT"
'openDB (C_TAB)
''Mengembalikan nilai Sebelumnya
'C_TAB = "INSERT INTO DAT_ZNT SELECT * FROM C_ZNT"
'openDB (C_TAB)
'    'Duplikasi Tabel DAT_FASILITAS_BANGUNAN
'C_TAB = "DELETE  FROM DAT_FASILITAS_BANGUNAN"
'openDB (C_TAB)
'    'Mengembalikan nilai Sebelumnya
'C_TAB = "INSERT INTO DAT_FASILITAS_BANGUNAN SELECT * FROM C_FASILITAS "
'openDB (C_TAB)
'    'Duplikasi Tabel DAT_SUBJEK_PAJAK_NJOPTKP
'C_TAB = "DELETE  FROM DAT_SUBJEK_PAJAK_NJOPTKP"
'openDB (C_TAB)
'    'Mengembalikan nilai Sebelumnya
'C_TAB = "INSERT INTO DAT_SUBJEK_PAJAK_NJOPTKP SELECT * FROM C_NJOPTKP"
'openDB (C_TAB)
'HAPUS_TAB
'ins_TAB
C_STR = "PULIH_DUPLIKASI"
openDB (C_STR)
MsgBox "Proses Pemulihan Data Berhasil...!", vbInformation, "SUKSES!"
    

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub

Private Sub mnResource_Click()
frmResource.Show
End Sub

Private Sub mnSearch_Click()
On Error Resume Next
'xID = 0
frmLIST_Objek1.Show
End Sub

Private Sub mnSet_Click()
On Error Resume Next
If mnSet.Checked = True Then
    cekTampil = 0
    mnSet.Checked = False
Else
    cekTampil = 1
    mnSet.Checked = True
End If
End Sub

Private Sub mnSetting_Click()
On Error Resume Next
Select Case J_CETAK
Case 1
    rSPPT.PrinterSetup Me.hWnd
Case 2
    rSTTS.PrinterSetup Me.hWnd
Case 3
    rDHKP.PrinterSetup Me.hWnd
Case 4
    rKelas1.PrinterSetup Me.hWnd
Case 5
    rSimulasi.PrinterSetup Me.hWnd
Case 6
    rSEM.PrinterSetup Me.hWnd
Case 10
    rINDIVIDU.PrinterSetup Me.hWnd
Case 20
    rDetail2.PrinterSetup Me.hWnd
Case 30
    rDETAIL.PrinterSetup Me.hWnd
Case 40
    rNilai.PrinterSetup Me.hWnd
Case 50
    rNilai2.PrinterSetup Me.hWnd
Case 60
    rNilai3.PrinterSetup Me.hWnd
Case 100
    rSPPT1.PrinterSetup Me.hWnd
Case 200
    rSSPD.PrinterSetup Me.hWnd
Case 300
    rDHKP1.PrinterSetup Me.hWnd
Case 400
    rSPPTt.PrinterSetup Me.hWnd
Case 500
    rSSPDt.PrinterSetup Me.hWnd
Case 111
    rLog1.PrinterSetup Me.hWnd
Case 112
    rLog2.PrinterSetup Me.hWnd
End Select
rptPBB.CRViewer1.Refresh
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'Salah:
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub MNSIMULASI_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 5
End Sub

Private Sub mnSistem3_8_Click()
On Error Resume Next
frmBentuk_DBKB.Show
cBentuk = 2
End Sub

Private Sub mnSPPT_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 1
End Sub

Private Sub mnSSH_Click()
On Error Resume Next
frmSSH.Show
End Sub

Private Sub mnSTTS_Click()
On Error Resume Next

Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 2
End Sub

Private Sub mnSubjek_Click()
On Error Resume Next
frmSubjek_Pajak.Show
End Sub

Private Sub mnSusut_Click()
On Error Resume Next
frmSusut.Show
End Sub

Private Sub mnTarif_Click()
On Error Resume Next
frmTarif.Show
End Sub

Private Sub mntunggal1_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 400

End Sub

Private Sub mntunggal2_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 500
End Sub

Private Sub mnUlin_Click()
frmUlin.Show
End Sub

Private Sub mnUser_Click()
On Error Resume Next
frmUser.Show
End Sub

Private Sub mnWewenang_Click()
On Error Resume Next
frmUser_Wewenang.Show
End Sub

Private Sub mnwSPPT_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 100
End Sub

Private Sub mnwSSPD_Click()
On Error Resume Next
Unload frmCetak_Massal
Unload rptPBB
frmCetak_Massal.Show
J_CETAK = 200
End Sub

Private Sub mnZNT1_Click()
On Error Resume Next
frmZNT.Show
End Sub

Private Sub nmJalan_Click()
On Error Resume Next
frmJalan.Show
End Sub

Sub CIPTA_TAB()
On Error GoTo Salah
'Duplikasi Tabel JPB2
C_TAB = "SELECT * INTO C2 FROM DAT_JPB2"
openDB (C_TAB)
'Duplikasi Tabel JPB3
C_TAB = "SELECT * INTO C3 FROM DAT_JPB3"
openDB (C_TAB)
'Duplikasi Tabel JPB4
C_TAB = "SELECT * INTO C4 FROM DAT_JPB4"
openDB (C_TAB)
'Duplikasi Tabel JPB5
C_TAB = "SELECT * INTO C5 FROM DAT_JPB5"
openDB (C_TAB)
'Duplikasi Tabel JPB6
C_TAB = "SELECT * INTO C6 FROM DAT_JPB6"
openDB (C_TAB)
'Duplikasi Tabel JPB7
C_TAB = "SELECT * INTO C7 FROM DAT_JPB7"
openDB (C_TAB)
'Duplikasi Tabel JPB8
C_TAB = "SELECT * INTO C8 FROM DAT_JPB8"
openDB (C_TAB)
'Duplikasi Tabel JPB9
C_TAB = "SELECT * INTO C9 FROM DAT_JPB9"
openDB (C_TAB)
'Duplikasi Tabel JPB12
C_TAB = "SELECT * INTO C12 FROM DAT_JPB12"
openDB (C_TAB)
'Duplikasi Tabel JPB13
C_TAB = "SELECT * INTO C13 FROM DAT_JPB13"
openDB (C_TAB)
'Duplikasi Tabel JPB14
C_TAB = "SELECT * INTO C14 FROM DAT_JPB14"
openDB (C_TAB)
'Duplikasi Tabel JPB15
C_TAB = "SELECT * INTO C15 FROM DAT_JPB15"
openDB (C_TAB)
'Duplikasi Tabel JPB16
C_TAB = "SELECT * INTO C16 FROM DAT_JPB16"
openDB (C_TAB)
'Duplikasi Tabel JPB17
C_TAB = "SELECT * INTO C17 FROM DAT_JPB17"
openDB (C_TAB)
Salah:
If Err.Number = 0 Then
    Exit Sub
ElseIf Err.Number = -2147217900 Then
    tt1 = MsgBox("Tabel sudah berisi, apa ingin ditimpa?", vbQuestion + vbYesNo, "Exist...!")
    If tt1 = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
    'Duplikasi Tabel JPB2
        C_TAB = "SELECT * INTO C2 FROM DAT_JPB2"
        openDB (C_TAB)
        'Duplikasi Tabel JPB3
        C_TAB = "SELECT * INTO C3 FROM DAT_JPB3"
        openDB (C_TAB)
        'Duplikasi Tabel JPB4
        C_TAB = "SELECT * INTO C4 FROM DAT_JPB4"
        openDB (C_TAB)
        'Duplikasi Tabel JPB5
        C_TAB = "SELECT * INTO C5 FROM DAT_JPB5"
        openDB (C_TAB)
        'Duplikasi Tabel JPB6
        C_TAB = "SELECT * INTO C6 FROM DAT_JPB6"
        openDB (C_TAB)
        'Duplikasi Tabel JPB7
        C_TAB = "SELECT * INTO C7 FROM DAT_JPB7"
        openDB (C_TAB)
        'Duplikasi Tabel JPB8
        C_TAB = "SELECT * INTO C8 FROM DAT_JPB8"
        openDB (C_TAB)
        'Duplikasi Tabel JPB9
        C_TAB = "SELECT * INTO C9 FROM DAT_JPB9"
        openDB (C_TAB)
        'Duplikasi Tabel JPB12
        C_TAB = "SELECT * INTO C12 FROM DAT_JPB12"
        openDB (C_TAB)
        'Duplikasi Tabel JPB13
        C_TAB = "SELECT * INTO C13 FROM DAT_JPB13"
        openDB (C_TAB)
        'Duplikasi Tabel JPB14
        C_TAB = "SELECT * INTO C14 FROM DAT_JPB14"
        openDB (C_TAB)
        'Duplikasi Tabel JPB15
        C_TAB = "SELECT * INTO C15 FROM DAT_JPB15"
        openDB (C_TAB)
        'Duplikasi Tabel JPB16
        C_TAB = "SELECT * INTO C16 FROM DAT_JPB16"
        openDB (C_TAB)
        'Duplikasi Tabel JPB17
        C_TAB = "SELECT * INTO C17 FROM DAT_JPB17"
        openDB (C_TAB)
    HAPUS_TAB
    MsgBox "Proses Duplikasi Berhasil...!", vbInformation, "SUKSES!"
Else
    MsgBox Err.Number & " : " & Err.Description, vbCritical, "Error"
End If
Screen.MousePointer = vbDefault
End Sub
Sub HAPUS_TAB()
On Error GoTo Salah
'Hapus Tabel JPB2
C_TAB = "DELETE  From  C2 "
openDB (C_TAB)
'Hapus Tabel JPB3
C_TAB = "DELETE  From  C3"
openDB (C_TAB)
'Hapus Tabel JPB4
C_TAB = "DELETE  From  C4"
openDB (C_TAB)
'Hapus Tabel JPB5
C_TAB = "DELETE  From  C5"
openDB (C_TAB)
'Hapus Tabel JPB6
C_TAB = "DELETE  From  C6"
openDB (C_TAB)
'Hapus Tabel JPB7
C_TAB = "DELETE  From  C7 "
openDB (C_TAB)
'Hapus Tabel JPB8
C_TAB = "DELETE  From  C8 "
openDB (C_TAB)
'Hapus Tabel JPB9
C_TAB = "DELETE  From  C9 "
openDB (C_TAB)
'Hapus Tabel JPB12
C_TAB = "DELETE  From  C12"
openDB (C_TAB)
'Hapus Tabel JPB13
C_TAB = "DELETE  From  C13 "
openDB (C_TAB)
'Hapus Tabel JPB14
C_TAB = "DELETE  From  C14 "
openDB (C_TAB)
'Hapus Tabel JPB15
C_TAB = "DELETE  From  C15 "
openDB (C_TAB)
'Hapus Tabel JPB16
C_TAB = "DELETE  From  C16 "
openDB (C_TAB)
'Hapus Tabel JPB17
C_TAB = "DELETE  From  C17 "
openDB (C_TAB)
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub ins_TAB()
On Error GoTo Salah
'Pemulihan Tabel JPB2
C_TAB = "INSERT  INTO DAT_JPB2 SELECT * FROM C2"
openDB (C_TAB)
'Pemulihan Tabel JPB3
C_TAB = "INSERT  INTO DAT_JPB3 SELECT * FROM C3"
openDB (C_TAB)
'Pemulihan Tabel JPB4
C_TAB = "INSERT  INTO DAT_JPB4 SELECT * FROM C4"
openDB (C_TAB)
'Pemulihan Tabel JPB5
C_TAB = "INSERT  INTO DAT_JPB5 SELECT * FROM C5"
openDB (C_TAB)
'Pemulihan Tabel JPB6
C_TAB = "INSERT  INTO DAT_JPB6 SELECT * FROM C6"
openDB (C_TAB)
'Pemulihan Tabel JPB7
C_TAB = "INSERT  INTO DAT_JPB7 SELECT * FROM C7"
openDB (C_TAB)
'Pemulihan Tabel JPB8
C_TAB = "INSERT  INTO DAT_JPB8 SELECT * FROM C8"
openDB (C_TAB)
'Pemulihan Tabel JPB9
C_TAB = "INSERT  INTO DAT_JPB9 SELECT * FROM C9"
openDB (C_TAB)
'Pemulihan Tabel JPB12
C_TAB = "INSERT  INTO DAT_JPB12 SELECT * FROM C12"
openDB (C_TAB)
'Pemulihan Tabel JPB13
C_TAB = "INSERT  INTO DAT_JPB13 SELECT * FROM C13"
openDB (C_TAB)
'Pemulihan Tabel JPB14
C_TAB = "INSERT  INTO DAT_JPB14 SELECT * FROM C14"
openDB (C_TAB)
'Pemulihan Tabel JPB15
C_TAB = "INSERT  INTO DAT_JPB15 SELECT * FROM C15"
openDB (C_TAB)
'Pemulihan Tabel JPB16
C_TAB = "INSERT  INTO DAT_JPB16 SELECT * FROM C16"
openDB (C_TAB)
'Pemulihan Tabel JPB17
C_TAB = "INSERT  INTO DAT_JPB17 SELECT * FROM C17"
openDB (C_TAB)
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
stsBar.Panels.Item(5).Text = Format(Time, "HH:MM:SS")
End Sub

