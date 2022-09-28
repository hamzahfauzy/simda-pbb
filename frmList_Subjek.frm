VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmList_Subjek 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List : Subjek Pajak"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11070
   Icon            =   "frmList_Subjek.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   75
      Picture         =   "frmList_Subjek.frx":1CCA
      ScaleHeight     =   300
      ScaleWidth      =   10920
      TabIndex        =   14
      Top             =   1050
      Width           =   10920
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1680
         TabIndex        =   16
         Top             =   75
         Width           =   555
      End
      Begin VB.Label LFIN 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Pencarian"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   4050
         TabIndex        =   15
         Top             =   75
         Width           =   6825
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   75
      Picture         =   "frmList_Subjek.frx":6332
      ScaleHeight     =   300
      ScaleWidth      =   10905
      TabIndex        =   12
      Top             =   180
      Width           =   10905
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Pencarian"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1140
         TabIndex        =   13
         Top             =   90
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5715
      TabIndex        =   9
      Top             =   7275
      Width           =   915
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4815
      TabIndex        =   10
      Top             =   7275
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3915
      TabIndex        =   11
      Top             =   7275
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   4965
      TabIndex        =   5
      Top             =   420
      Width           =   6045
      Begin VB.TextBox tCari 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1245
         TabIndex        =   7
         Top             =   195
         Width           =   4305
      End
      Begin VB.CommandButton cmdCari 
         Height          =   375
         Left            =   5520
         Picture         =   "frmList_Subjek.frx":A99A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label1Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ketik Teks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   255
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   75
      TabIndex        =   0
      Top             =   420
      Width           =   4860
      Begin VB.OptionButton oKlas 
         Caption         =   "No. Identitas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   150
         Picture         =   "frmList_Subjek.frx":ACF5
         TabIndex        =   3
         Top             =   300
         Width           =   1575
      End
      Begin VB.OptionButton oKlas 
         Caption         =   "Nama WP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1860
         Picture         =   "frmList_Subjek.frx":191DE
         TabIndex        =   2
         Top             =   300
         Width           =   1440
      End
      Begin VB.OptionButton oKlas 
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3435
         Picture         =   "frmList_Subjek.frx":276C7
         TabIndex        =   1
         Top             =   315
         Width           =   1245
      End
   End
   Begin MSComctlLib.ListView vArsip 
      Height          =   5730
      Left            =   60
      TabIndex        =   4
      Top             =   1365
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   10107
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483642
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "No"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "b"
         Text            =   "No. ID"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nama Wajib Pajak"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Jalan/Dusun"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Kelurahan/Desa"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Kota/Kab"
         Object.Width           =   6774
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "NPWP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "BLOK"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "RW"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "RT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "POS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PEKERJAAN"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   9195
      Left            =   -30
      Picture         =   "frmList_Subjek.frx":35BB0
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   12960
   End
End
Attribute VB_Name = "frmList_Subjek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pilih
Private Sub cmdCari_Click()
On Error GoTo Salah
vArsip.Sorted = False
    vArsip.Sorted = False
Screen.MousePointer = vbHourglass
    Set rPajak = Nothing
    
    vArsip.ListItems.Clear
    
    If Pilih = 1 Then
    SQLStr1 = "select * from   DAT_SUBJEK_PAJAK WHERE SUBJEK_PAJAK_ID LIKE '" & "%" & Trim(tCari.Text) & "%" & "' order by SUBJEK_PAJAK_ID asc"
    openDB (SQLStr1)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
    i = i + 1
    vArsip.ListItems.Add i, "", Format(i, "#")
    vArsip.ListItems.Item(i).ListSubItems.Add 1, "", i
    xxID = Trim(rPajak![SUBJEK_PAJAK_ID])
    If IsNull(Trim(rPajak![SUBJEK_PAJAK_ID])) = True Or Trim(rPajak![SUBJEK_PAJAK_ID]) = "" Then xxID = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 2, "", xxID 'Trim(rPajak![SUBJEK_PAJAK_ID])
    xxNAMA = Trim(rPajak![Nm_wp])
    If IsNull(Trim(rPajak![Nm_wp])) = True Or Trim(rPajak![Nm_wp]) = "" Then xxNAMA = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 3, "", xxNAMA 'Trim(rPajak![NM_WP])
    xxJalan = Trim(rPajak![JALAN_WP])
    If IsNull(Trim(rPajak![JALAN_WP])) = True Or Trim(rPajak![JALAN_WP]) = "" Then xxJalan = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 4, "", xxJalan 'Trim(rPajak![JALAN_WP])
    xKel = Trim(rPajak![KELURAHAN_WP])
    If IsNull(Trim(rPajak![KELURAHAN_WP])) = True Or Trim(rPajak![KELURAHAN_WP]) = "" Then xKel = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 5, "", xKel 'Trim(rPajak![KELURAHAN_WP])
   ' vArsip.ListItems.Item(I).ListSubItems.Add 6, "", Trim(rPajak![KOTA_WP])
   
   ' vArsip.ListItems.Item(I).ListSubItems.Add 7, "", Trim(rPajak![NPWP])
   xKota = Trim(rPajak![KOTA_WP])
    If IsNull(Trim(rPajak![KOTA_WP])) = True Or Trim(rPajak![KOTA_WP]) = "" Then xKota = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 6, "", xKota 'Trim(rPajak![KOTA_WP])
    xNPWP = Trim(rPajak![NPWP])
     If IsNull(Trim(rPajak![NPWP])) = True Or Trim(rPajak![NPWP]) = "" Then xNPWP = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 7, "", xNPWP 'Trim(rPajak![NPWP])
    xBLOK = Trim(rPajak![BLOK_KAV_NO_WP])
    If IsNull(Trim(rPajak![BLOK_KAV_NO_WP])) = True Or Trim(rPajak![BLOK_KAV_NO_WP]) = "" Then xBLOK = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 8, "", xBLOK 'Trim(rPajak![BLOK_KAV_NO_WP])
   xRW = Trim(rPajak![RW_WP])
    If IsNull(Trim(rPajak![RW_WP])) = True Or Trim(rPajak![RW_WP]) = "" Then xRW = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 9, "", xRW 'Trim(rPajak![RW_WP])
  xRT = Trim(rPajak![RT_WP])
    If IsNull(Trim(rPajak![RT_WP])) = True Or Trim(rPajak![RT_WP]) = "" Then xRT = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 10, "", xRT 'Trim(rPajak![RT_WP])
  xPos = Trim(rPajak![KD_POS_WP])
    If IsNull(Trim(rPajak![KD_POS_WP])) = True Or Trim(rPajak![KD_POS_WP]) = "" Then xPos = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 11, "", xPos 'Trim(rPajak![KD_POS_WP])
  xKerja = Trim(rPajak![STATUS_PEKERJAAN_WP])
    If IsNull(Trim(rPajak![STATUS_PEKERJAAN_WP])) = True Or Trim(rPajak![STATUS_PEKERJAAN_WP]) = "" Then xKerja = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 12, "", xKerja 'Trim(rPajak![STATUS_PEKERJAAN_WP])
    rPajak.MoveNext
    Loop
    '----------------------
    ElseIf Pilih = 2 Then
    SQLStr2 = "select * from   DAT_SUBJEK_PAJAK WHERE NM_WP LIKE '" & "%" & Trim(tCari.Text) & "%" & "' order by NM_WP,SUBJEK_PAJAK_ID asc"
    openDB (SQLStr2)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
    i = i + 1
   vArsip.ListItems.Add i, "", Format(i, "#")
    vArsip.ListItems.Item(i).ListSubItems.Add 1, "", i
  xxID = Trim(rPajak![SUBJEK_PAJAK_ID])
    If IsNull(Trim(rPajak![SUBJEK_PAJAK_ID])) = True Or Trim(rPajak![SUBJEK_PAJAK_ID]) = "" Then xxID = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 2, "", xxID 'Trim(rPajak![SUBJEK_PAJAK_ID])
    xxNAMA = Trim(rPajak![Nm_wp])
    If IsNull(Trim(rPajak![Nm_wp])) = True Or Trim(rPajak![Nm_wp]) = "" Then xxNAMA = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 3, "", xxNAMA 'Trim(rPajak![NM_WP])
    xxJalan = Trim(rPajak![JALAN_WP])
    If IsNull(Trim(rPajak![JALAN_WP])) = True Or Trim(rPajak![JALAN_WP]) = "" Then xxJalan = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 4, "", xxJalan 'Trim(rPajak![JALAN_WP])
    xKel = Trim(rPajak![KELURAHAN_WP])
    If IsNull(Trim(rPajak![KELURAHAN_WP])) = True Or Trim(rPajak![KELURAHAN_WP]) = "" Then xKel = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 5, "", xKel 'Trim(rPajak![KELURAHAN_WP])
    xKota = Trim(rPajak![KOTA_WP])
    If IsNull(Trim(rPajak![KOTA_WP])) = True Or Trim(rPajak![KOTA_WP]) = "" Then xKota = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 6, "", xKota 'Trim(rPajak![KOTA_WP])
    xNPWP = Trim(rPajak![NPWP])
     If IsNull(Trim(rPajak![NPWP])) = True Or Trim(rPajak![NPWP]) = "" Then xNPWP = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 7, "", xNPWP 'Trim(rPajak![NPWP])
    xBLOK = Trim(rPajak![BLOK_KAV_NO_WP])
    If IsNull(Trim(rPajak![BLOK_KAV_NO_WP])) = True Or Trim(rPajak![BLOK_KAV_NO_WP]) = "" Then xBLOK = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 8, "", xBLOK 'Trim(rPajak![BLOK_KAV_NO_WP])
   xRW = Trim(rPajak![RW_WP])
    If IsNull(Trim(rPajak![RW_WP])) = True Or Trim(rPajak![RW_WP]) = "" Then xRW = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 9, "", xRW 'Trim(rPajak![RW_WP])
  xRT = Trim(rPajak![RT_WP])
    If IsNull(Trim(rPajak![RT_WP])) = True Or Trim(rPajak![RT_WP]) = "" Then xRT = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 10, "", xRT 'Trim(rPajak![RT_WP])
  xPos = Trim(rPajak![KD_POS_WP])
    If IsNull(Trim(rPajak![KD_POS_WP])) = True Or Trim(rPajak![KD_POS_WP]) = "" Then xPos = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 11, "", xPos 'Trim(rPajak![KD_POS_WP])
  xKerja = Trim(rPajak![STATUS_PEKERJAAN_WP])
    If IsNull(Trim(rPajak![STATUS_PEKERJAAN_WP])) = True Or Trim(rPajak![STATUS_PEKERJAAN_WP]) = "" Then xKerja = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 12, "", xKerja 'Trim(rPajak![STATUS_PEKERJAAN_WP])
    rPajak.MoveNext
    Loop
    '-------------------
    ElseIf Pilih = 3 Then
    SQLStr3 = "select * from   DAT_SUBJEK_PAJAK WHERE JALAN_WP +' ' + KELURAHAN_WP +' ' + KOTA_WP LIKE '" & "%" & Trim(tCari.Text) & "%" & "' order by JALAN_WP,KELURAHAN_WP,SUBJEK_PAJAK_ID asc"
    openDB (SQLStr3)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
    i = i + 1
   vArsip.ListItems.Add i, "", Format(i, "#")
    vArsip.ListItems.Item(i).ListSubItems.Add 1, "", i
   xxID = Trim(rPajak![SUBJEK_PAJAK_ID])
    If IsNull(Trim(rPajak![SUBJEK_PAJAK_ID])) = True Or Trim(rPajak![SUBJEK_PAJAK_ID]) = "" Then xxID = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 2, "", xxID 'Trim(rPajak![SUBJEK_PAJAK_ID])
    xxNAMA = Trim(rPajak![Nm_wp])
    If IsNull(Trim(rPajak![Nm_wp])) = True Or Trim(rPajak![Nm_wp]) = "" Then xxNAMA = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 3, "", xxNAMA 'Trim(rPajak![NM_WP])
    xxJalan = Trim(rPajak![JALAN_WP])
    If IsNull(Trim(rPajak![JALAN_WP])) = True Or Trim(rPajak![JALAN_WP]) = "" Then xxJalan = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 4, "", xxJalan 'Trim(rPajak![JALAN_WP])
    xKel = Trim(rPajak![KELURAHAN_WP])
    If IsNull(Trim(rPajak![KELURAHAN_WP])) = True Or Trim(rPajak![KELURAHAN_WP]) = "" Then xKel = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 5, "", xKel 'Trim(rPajak![KELURAHAN_WP])
    xKota = Trim(rPajak![KOTA_WP])
    If IsNull(Trim(rPajak![KOTA_WP])) = True Or Trim(rPajak![KOTA_WP]) = "" Then xKota = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 6, "", xKota 'Trim(rPajak![KOTA_WP])
    xNPWP = Trim(rPajak![NPWP])
     If IsNull(Trim(rPajak![NPWP])) = True Or Trim(rPajak![NPWP]) = "" Then xNPWP = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 7, "", xNPWP 'Trim(rPajak![NPWP])
    xBLOK = Trim(rPajak![BLOK_KAV_NO_WP])
    If IsNull(Trim(rPajak![BLOK_KAV_NO_WP])) = True Or Trim(rPajak![BLOK_KAV_NO_WP]) = "" Then xBLOK = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 8, "", xBLOK 'Trim(rPajak![BLOK_KAV_NO_WP])
   xRW = Trim(rPajak![RW_WP])
    If IsNull(Trim(rPajak![RW_WP])) = True Or Trim(rPajak![RW_WP]) = "" Then xRW = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 9, "", xRW 'Trim(rPajak![RW_WP])
  xRT = Trim(rPajak![RT_WP])
    If IsNull(Trim(rPajak![RT_WP])) = True Or Trim(rPajak![RT_WP]) = "" Then xRT = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 10, "", xRT 'Trim(rPajak![RT_WP])
  xPos = Trim(rPajak![KD_POS_WP])
    If IsNull(Trim(rPajak![KD_POS_WP])) = True Or Trim(rPajak![KD_POS_WP]) = "" Then xPos = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 11, "", xPos 'Trim(rPajak![KD_POS_WP])
  xKerja = Trim(rPajak![STATUS_PEKERJAAN_WP])
    If IsNull(Trim(rPajak![STATUS_PEKERJAAN_WP])) = True Or Trim(rPajak![STATUS_PEKERJAAN_WP]) = "" Then xKerja = "-"
    vArsip.ListItems.Item(i).ListSubItems.Add 12, "", xKerja 'Trim(rPajak![STATUS_PEKERJAAN_WP])
    rPajak.MoveNext
    Loop
    '----------------------
    
    End If
LFIN.Caption = ""
LFIN.Caption = "Jumlah " & oKlas(Pilih - 1).Caption & " : " & vArsip.ListItems.Count & " File"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExit_Click()
Unload Me
'If xID = 2 Then
'        frmSubjek_Pajak.tID(0).Text = ""
'End If
xID = 0
End Sub

Private Sub cmdOK_Click()
'On Error Resume Next
On Error GoTo Salah
Dim xKerja
If xID = 1 Then
    frmObjek_Pajak_Bm.tID1.Text = vArsip.SelectedItem.ListSubItems(2).Text
    'frmObjek_Pajak_Bm.tID2.Text = "Nama" & vbTab & ": " & vArsip.SelectedItem.ListSubItems(3).Text & vbCrLf & "Alamat" & vbTab & ": " & vArsip.SelectedItem.ListSubItems(4).Text & ", " & vArsip.SelectedItem.ListSubItems(5).Text & ", " & vArsip.SelectedItem.ListSubItems(6).Text & ", " & vArsip.SelectedItem.ListSubItems(7).Text
    'frmObjek_Pajak_Bm.LAlamat.Caption = vArsip.SelectedItem.ListSubItems(4).Text & ", " & vArsip.SelectedItem.ListSubItems(5).Text & "-" & vArsip.SelectedItem.ListSubItems(6).Text
    Unload Me
ElseIf xID = 2 Then
    frmSubjek_Pajak.tID(0).Text = vArsip.SelectedItem.ListSubItems(2).Text
    frmSubjek_Pajak.LID.Caption = vArsip.SelectedItem.ListSubItems(2).Text
    frmSubjek_Pajak.tID(1).Text = vArsip.SelectedItem.ListSubItems(3).Text ' & vbCrLf & "Alamat" & vbTab & ": " & vArsip.SelectedItem.ListSubItems(4).Text & ", " & vArsip.SelectedItem.ListSubItems(5).Text & ", " & vArsip.SelectedItem.ListSubItems(6).Text & ", " & vArsip.SelectedItem.ListSubItems(7).Text
    If vArsip.SelectedItem.ListSubItems(7).Text = "" Or IsNull(vArsip.SelectedItem.ListSubItems(7).Text) = True Then
        frmSubjek_Pajak.tID(2).Text = "-"
    Else
        frmSubjek_Pajak.tID(2).Text = vArsip.SelectedItem.ListSubItems(7).Text
    End If
    frmSubjek_Pajak.tID(3).Text = vArsip.SelectedItem.ListSubItems(4).Text
    frmSubjek_Pajak.tID(4).Text = vArsip.SelectedItem.ListSubItems(5).Text
    frmSubjek_Pajak.tID(8).Text = vArsip.SelectedItem.ListSubItems(6).Text
    If vArsip.SelectedItem.ListSubItems(8).Text = "" Or IsNull(vArsip.SelectedItem.ListSubItems(8).Text) = True Then
        frmSubjek_Pajak.tID(5).Text = "-"
    Else
        frmSubjek_Pajak.tID(5).Text = vArsip.SelectedItem.ListSubItems(8).Text
    End If
    'frmSubjek_Pajak.tID(5).Text = vArsip.SelectedItem.ListSubItems(7).Text
    If vArsip.SelectedItem.ListSubItems(9).Text = "" Or IsNull(vArsip.SelectedItem.ListSubItems(9).Text) = True Then
        frmSubjek_Pajak.tID(6).Text = "-"
    Else
        frmSubjek_Pajak.tID(6).Text = vArsip.SelectedItem.ListSubItems(9).Text
    End If
    'frmSubjek_Pajak.tID(6).Text = vArsip.SelectedItem.ListSubItems(8).Text
    If vArsip.SelectedItem.ListSubItems(10).Text = "" Or IsNull(vArsip.SelectedItem.ListSubItems(10).Text) = True Then
        frmSubjek_Pajak.tID(7).Text = "-"
    Else
        frmSubjek_Pajak.tID(7).Text = vArsip.SelectedItem.ListSubItems(10).Text
    End If
    'frmSubjek_Pajak.tID(7).Text = vArsip.SelectedItem.ListSubItems(9).Text
    If vArsip.SelectedItem.ListSubItems(11).Text = "" Or IsNull(vArsip.SelectedItem.ListSubItems(11).Text) = True Then
        frmSubjek_Pajak.tID(9).Text = "-"
    Else
        frmSubjek_Pajak.tID(9).Text = vArsip.SelectedItem.ListSubItems(11).Text
    End If
    'frmSubjek_Pajak.tID(9).Text = vArsip.SelectedItem.ListSubItems(11).Text
    
    xKerja = vArsip.SelectedItem.ListSubItems(12).Text '* 1
    If xKerja = 0 Or xKerja = "" Or IsNull(xKerja) = True Then
        frmSubjek_Pajak.ccKerja.Text = frmSubjek_Pajak.ccKerja.List(frmSubjek_Pajak.ccKerja.ListCount - 1)
    Else
        frmSubjek_Pajak.ccKerja.Text = frmSubjek_Pajak.ccKerja.List((xKerja) - 1)
    End If
    Unload Me
Else

    Unload Me
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub Form_Activate()
On Error Resume Next
Screen.MousePointer = vbHourglass
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
Pilih = 2
oKlas(1).SetFocus

tCari.SetFocus

Screen.MousePointer = vbDefault
End Sub

Private Sub tCari_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'If InStr("'", Chr(KeyAscii)) <> 0 And KeyAscii <> vbKeyBack Then
'            KeyAscii = 0
'        End If

If KeyAscii = 13 Then
    KeyAscii = 0
   tCari.Text = Rep(tCari.Text)
   cmdCari_Click
   
End If



End Sub


Private Sub tCari_LostFocus()
   
tCari.Text = Rep(tCari.Text)
End Sub

Private Sub vArsip_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vArsip.SortKey = ColumnHeader.Index - 1
vArsip.Sorted = True
vArsip.Sorted = False
vArsip.SortOrder = lvwAscending

End Sub

Private Sub oKlas_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
    Pilih = 1
Case 1
    Pilih = 2
Case 2
    Pilih = 3
Case 3
    Pilih = 4
Case 4
    Pilih = 5
End Select
tCari.SetFocus
tCari.Text = ""
End Sub
