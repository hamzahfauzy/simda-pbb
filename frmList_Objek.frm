VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmList_Objek 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List : Objek Pajak"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11085
   Icon            =   "frmList_Objek.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   60
      Picture         =   "frmList_Objek.frx":1CCA
      ScaleHeight     =   300
      ScaleWidth      =   10905
      TabIndex        =   17
      Top             =   90
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
         TabIndex        =   18
         Top             =   90
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   60
      Picture         =   "frmList_Objek.frx":6332
      ScaleHeight     =   300
      ScaleWidth      =   10920
      TabIndex        =   13
      Top             =   930
      Width           =   10920
      Begin VB.Label Label3 
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
         Left            =   1530
         TabIndex        =   19
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
         Height          =   180
         Left            =   4095
         TabIndex        =   14
         Top             =   90
         Width           =   6795
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
      Left            =   5700
      TabIndex        =   8
      Top             =   7740
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
      Left            =   4800
      TabIndex        =   9
      Top             =   7740
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
      Left            =   3900
      TabIndex        =   10
      Top             =   7740
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   4935
      TabIndex        =   4
      Top             =   300
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
         Left            =   1215
         TabIndex        =   6
         Top             =   195
         Width           =   4305
      End
      Begin VB.CommandButton cmdCari 
         Height          =   375
         Left            =   5520
         Picture         =   "frmList_Objek.frx":A99A
         Style           =   1  'Graphical
         TabIndex        =   5
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
         TabIndex        =   7
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
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   4860
      Begin VB.OptionButton oKlas 
         Caption         =   "Lokasi Objek Pajak"
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
         Left            =   120
         Picture         =   "frmList_Objek.frx":ACF5
         TabIndex        =   11
         Top             =   285
         Width           =   1650
      End
      Begin VB.OptionButton oKlas 
         Caption         =   "N.O.P"
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
         Left            =   2040
         Picture         =   "frmList_Objek.frx":191DE
         TabIndex        =   2
         Top             =   300
         Width           =   1125
      End
      Begin VB.OptionButton oKlas 
         Caption         =   "No. Formulir"
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
         Left            =   3465
         Picture         =   "frmList_Objek.frx":276C7
         TabIndex        =   1
         Top             =   315
         Width           =   1320
      End
   End
   Begin MSComctlLib.ListView vBumi 
      Height          =   6450
      Left            =   45
      TabIndex        =   3
      Top             =   1245
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   11377
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
      NumItems        =   32
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Nomor Objek Pajak (NOP)"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Lokasi Objek Pajak"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   "b"
         Text            =   "Kode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Desa/Kelurahan"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Kode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Kecamatan"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Luas"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Nilai Sistem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Kode"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Jenis Tanah"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Formulir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "ID WP"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Status WP"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "No Persil"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "ZNT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "NM Jalan OP"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Kav OP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "RW OP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "RT OP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "TGL Data"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "NIP Pendata"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Tgl Periksa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "NIP Periksa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Tgl Rekam"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "NIP Perekam"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "5"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView vBangunan 
      Height          =   1875
      Left            =   30
      TabIndex        =   12
      Top             =   5805
      Visible         =   0   'False
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   3307
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "KODE_JPB"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jenis Penggunaan Bangunan (JPB)"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Luas"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Nilai Sistem"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "BNG_Ke"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Formulir LSPOP"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   45
      Picture         =   "frmList_Objek.frx":35BB0
      ScaleHeight     =   300
      ScaleWidth      =   10920
      TabIndex        =   15
      Top             =   5490
      Visible         =   0   'False
      Width           =   10920
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBJEK PAJAK BANGUNAN"
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
         Left            =   855
         TabIndex        =   16
         Top             =   90
         Width           =   2205
      End
   End
   Begin VB.Image Image1 
      Height          =   9195
      Left            =   -30
      Picture         =   "frmList_Objek.frx":3A218
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   12960
   End
End
Attribute VB_Name = "frmList_Objek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pilih
Private Sub cmdCari_Click()
On Error GoTo Salah
Dim xxKec, xxKel, xxLuas, xxNilai, xxForm, xxSektor, xxNSektor, xxMAP, xxZNT
vBumi.Sorted = False
    vBumi.Sorted = False
    'Screen.MousePointer = vbHourglass
    Set rPajak = Nothing
    
    vBumi.ListItems.Clear
    
    If Pilih = 1 Then
    'StringQ = "SELECT DAT_OP_BUMI.NO_FORMULIR,DAT_OP_BUMI.JNS_BUMI,DAT_OP_BUMI.LUAS_BUMI,DAT_OP_BUMI.NILAI_SISTEM_BUMI, DAT_OP_BUMI.KD_PROPINSI, DAT_OP_BUMI.KD_DATI2, DAT_OP_BUMI.KD_Kecamatan, DAT_OP_BUMI.KD_KELURAHAN, DAT_OP_BUMI.KD_BLOK, DAT_OP_BUMI.NO_URUT, DAT_OP_BUMI.KD_JNS_OP, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, " & _
     vbCrLf & "[DAT_OP_BUMI].[KD_PROPINSI] + '.' + [DAT_OP_BUMI].[KD_DATI2] + '.' + [DAT_OP_BUMI].[KD_KECAMATAN] + '.' + [DAT_OP_BUMI].[KD_KELURAHAN] + '.' + [DAT_OP_BUMI].[KD_BLOK] + '-' + [DAT_OP_BUMI].[NO_URUT] + '.' + [DAT_OP_BUMI].[KD_JNS_OP] AS NOPQ " & _
    vbCrLf & "FROM (DAT_OP_BUMI INNER JOIN REF_KECAMATAN ON DAT_OP_BUMI.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) INNER JOIN REF_KELURAHAN ON (REF_KELURAHAN.KD_KECAMATAN = DAT_OP_BUMI.KD_KECAMATAN) AND (DAT_OP_BUMI.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) Where REF_KELURAHAN.NM_KELURAHAN LIKE '" & "%" & Trim(tCari.Text) & "%" & "'"
    StringQ = "select * from QOBJEKPAJAK where [NM_KELURAHAN] + ', KEC. ' + [NM_KECAMATAN] LIKE '" & "%" & Trim(tCari.Text) & "%" & "' order by NOPQ asc"
    'StringQ = "select * from vPBB where NM_KELURAHAN LIKE '" & "%" & Trim(tCari.Text) & "%" & "'"
    openDB (StringQ)
    
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    I = 0
'    Do While Not rPajak.EOF
'    'If rPajak!NM_KELURAHAN Like "*" & Trim(tCari.Text) & "*" Then
'    I = I + 1
'    vBumi.ListItems.Add I, "", Format(I, "#")
'    vBumi.ListItems.Item(I).ListSubItems.Add 1, "", Format(I, "#")
'    'NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP)
'            vBumi.ListItems.Item(I).ListSubItems.Add 2, "", rPajak!NOPQ
'            xxKel = Trim(rPajak![NM_KELURAHAN])
'            If xxKel = True Or xxKel = "" Then xxKel = "-"
'            xxKec = Trim(rPajak![NM_KECAMATAN])
'            If xxKec = True Or xxKec = "" Then xxKec = "-"
'            vBumi.ListItems.Item(I).ListSubItems.Add 3, "", xxKel & ", KEC. " & xxKec 'Trim(rPajak![NM_KELURAHAN]) & ", KEC. " & Trim(rPajak![NM_KECAMATAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 4, "", xxKel 'Trim(rPajak![KD_KELURAHAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 5, "", xxKel 'Trim(rPajak![NM_KELURAHAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 6, "", xxKec 'Trim(rPajak![KD_KECAMATAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 7, "", xxKec 'Trim(rPajak![NM_KECAMATAN])
'            xxLuas = Trim(rPajak![LUAS_BUMI])
'            If IsNull(xxLuas) = True Or xxLuas = "" Then xxLuas = 0
'            vBumi.ListItems.Item(I).ListSubItems.Add 8, "", Format(xxLuas, "#,#0") 'Format(Trim(rPajak![LUAS_BUMI]), "#,#0")
'            xxNilai = Trim(rPajak![NILAI_SISTEM_BUMI])
'            If IsNull(xxNilai) = True Or xxNilai = "" Then xxNilai = 0
'            vBumi.ListItems.Item(I).ListSubItems.Add 9, "", Format(xxNilai, "#,#0") 'Format(Trim(rPajak![NILAI_SISTEM_BUMI]), "#,#0")
'            vBumi.ListItems.Item(I).ListSubItems.Add 10, "", Trim(rPajak![JNS_BUMI])
'            If Trim(rPajak![JNS_BUMI]) = 1 Then
'                JTANAH = "TANAH DAN BANGUNAN"
'            ElseIf Trim(rPajak![JNS_BUMI]) = 2 Then
'                JTANAH = "KAVLING DAN SIAP BANGUN"
'            ElseIf Trim(rPajak![JNS_BUMI]) = 3 Then
'                JTANAH = "TANAH KOSONG"
'            Else
'                JTANAH = "LAINNYA"
'            End If
'            vBumi.ListItems.Item(I).ListSubItems.Add 11, "", JTANAH
'            xxForm = Trim(rPajak![NO_FORMULIR])
'            If IsNull(xxForm) = True Or Trim(rPajak![NO_FORMULIR]) = "" Then xxForm = "-"
'            vBumi.ListItems.Item(I).ListSubItems.Add 12, "", xxForm 'Trim(rPajak![NO_FORMULIR])
'            xxSektor = Trim(rPajak![KD_SEKTOR])
'            If xxSektor = "" Or IsNull(xxSektor) = True Then xxSektor = "00"
'            vBumi.ListItems.Item(I).ListSubItems.Add 13, "", xxSektor 'Trim(rPajak![KD_SEKTOR])
'            If xxNSektor = "" Or IsNull(xxNSektor) = True Then xxNSektor = "-"
'            vBumi.ListItems.Item(I).ListSubItems.Add 14, "", xxNSektor 'Trim(rPajak![NM_SEKTOR])
'            If xxMAP = "" Or IsNull(xxMAP) = True Then xxMAP = "00000"
'            vBumi.ListItems.Item(I).ListSubItems.Add 15, "", xxMAP 'Trim(rPajak![KD_MAP])
'    'End If
'    rPajak.MoveNext
'    Loop
    '----------------------
    ElseIf Pilih = 2 Then
    'StringQ = "SELECT DAT_OP_BUMI.NO_FORMULIR,DAT_OP_BUMI.JNS_BUMI,DAT_OP_BUMI.LUAS_BUMI,DAT_OP_BUMI.NILAI_SISTEM_BUMI, DAT_OP_BUMI.KD_PROPINSI, DAT_OP_BUMI.KD_DATI2, DAT_OP_BUMI.KD_Kecamatan, DAT_OP_BUMI.KD_KELURAHAN, DAT_OP_BUMI.KD_BLOK, DAT_OP_BUMI.NO_URUT, DAT_OP_BUMI.KD_JNS_OP, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, " & _
     vbCrLf & "[DAT_OP_BUMI].[KD_PROPINSI] + '.' + [DAT_OP_BUMI].[KD_DATI2] + '.' + [DAT_OP_BUMI].[KD_KECAMATAN] + '.' + [DAT_OP_BUMI].[KD_KELURAHAN] + '.' + [DAT_OP_BUMI].[KD_BLOK] + '-' + [DAT_OP_BUMI].[NO_URUT] + '.' + [DAT_OP_BUMI].[KD_JNS_OP] AS NOPQ " & _
    vbCrLf & "FROM (DAT_OP_BUMI INNER JOIN REF_KECAMATAN ON DAT_OP_BUMI.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) INNER JOIN REF_KELURAHAN ON (REF_KELURAHAN.KD_KECAMATAN = DAT_OP_BUMI.KD_KECAMATAN) AND (DAT_OP_BUMI.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) Where DAT_OP_BUMI.LUAS_BUMI LIKE '" & "%" & Trim(tCari.Text) & "%" & "'"
    StringQ = "Select * From QOBJEKPAJAK where NOPQ LIKE '" & "%" & Trim(tCari.Text) & "%" & "' order by NOPQ asc"
    openDB (StringQ)
    
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    I = 0
'    Do While Not rPajak.EOF
'    'If rPajak!NOPQ Like "& % & Trim(tCari.Text) & % &" = True Then
'    I = I + 1
'    vBumi.ListItems.Add I, "", Format(I, "#")
'    vBumi.ListItems.Item(I).ListSubItems.Add 1, "", Format(I, "#")
'    'NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP)
'            vBumi.ListItems.Item(I).ListSubItems.Add 2, "", rPajak!NOPQ
'            vBumi.ListItems.Item(I).ListSubItems.Add 3, "", Trim(rPajak![NM_KELURAHAN]) & ", KEC. " & Trim(rPajak![NM_KECAMATAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 4, "", Trim(rPajak![KD_KELURAHAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 5, "", Trim(rPajak![NM_KELURAHAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 6, "", Trim(rPajak![KD_KECAMATAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 7, "", Trim(rPajak![NM_KECAMATAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 8, "", Format(Trim(rPajak![LUAS_BUMI]), "#,#0")
'            vBumi.ListItems.Item(I).ListSubItems.Add 9, "", Format(Trim(rPajak![NILAI_SISTEM_BUMI]), "#,#0")
'            vBumi.ListItems.Item(I).ListSubItems.Add 10, "", Trim(rPajak![JNS_BUMI])
'            If Trim(rPajak![JNS_BUMI]) = 1 Then
'                JTANAH = "TANAH DAN BANGUNAN"
'            ElseIf Trim(rPajak![JNS_BUMI]) = 2 Then
'                JTANAH = "KAVLING DAN SIAP BANGUN"
'            ElseIf Trim(rPajak![JNS_BUMI]) = 3 Then
'                JTANAH = "TANAH KOSONG"
'            Else
'                JTANAH = "LAINNYA"
'            End If
'            vBumi.ListItems.Item(I).ListSubItems.Add 11, "", JTANAH
'            vBumi.ListItems.Item(I).ListSubItems.Add 12, "", Trim(rPajak![NO_FORMULIR])
'            vBumi.ListItems.Item(I).ListSubItems.Add 13, "", Trim(rPajak![KD_SEKTOR])
'            vBumi.ListItems.Item(I).ListSubItems.Add 14, "", Trim(rPajak![NM_SEKTOR])
'            vBumi.ListItems.Item(I).ListSubItems.Add 15, "", Trim(rPajak![KD_MAP])
'    rPajak.MoveNext
'    Loop
    '-------------------
    ElseIf Pilih = 3 Then
    'stringQ = "SELECT DAT_OP_BUMI.NO_FORMULIR,DAT_OP_BUMI.JNS_BUMI,DAT_OP_BUMI.LUAS_BUMI,DAT_OP_BUMI.NILAI_SISTEM_BUMI, DAT_OP_BUMI.KD_PROPINSI, DAT_OP_BUMI.KD_DATI2, DAT_OP_BUMI.KD_Kecamatan, DAT_OP_BUMI.KD_KELURAHAN, DAT_OP_BUMI.KD_BLOK, DAT_OP_BUMI.NO_URUT, DAT_OP_BUMI.KD_JNS_OP, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, " & _
     vbCrLf & "[DAT_OP_BUMI].[KD_PROPINSI] + '.' + [DAT_OP_BUMI].[KD_DATI2] + '.' + [DAT_OP_BUMI].[KD_KECAMATAN] + '.' + [DAT_OP_BUMI].[KD_KELURAHAN] + '.' + [DAT_OP_BUMI].[KD_BLOK] + '-' + [DAT_OP_BUMI].[NO_URUT] + '.' + [DAT_OP_BUMI].[KD_JNS_OP] AS NOPQ " & _
    vbCrLf & "FROM (DAT_OP_BUMI INNER JOIN REF_KECAMATAN ON DAT_OP_BUMI.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) INNER JOIN REF_KELURAHAN ON (REF_KELURAHAN.KD_KECAMATAN = DAT_OP_BUMI.KD_KECAMATAN) AND (DAT_OP_BUMI.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) Where DAT_OP_BUMI.NO_FORMULIR LIKE '" & "%" & Trim(tCari.Text) & "%" & "'"
    StringQ = "select * from QOBJEKPAJAK where NO_FORMULIR_SPOP LIKE '" & "%" & Trim(tCari.Text) & "%" & "' order by NOPQ asc"
    openDB (StringQ)
    
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    I = 0
'    Do While Not rPajak.EOF
'    I = I + 1
'    vBumi.ListItems.Add I, "", Format(I, "#")
'    vBumi.ListItems.Item(I).ListSubItems.Add 1, "", Format(I, "#")
'    'NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP)
'            vBumi.ListItems.Item(I).ListSubItems.Add 2, "", rPajak!NOPQ
'            vBumi.ListItems.Item(I).ListSubItems.Add 3, "", Trim(rPajak![NM_KELURAHAN]) & ", KEC. " & Trim(rPajak![NM_KECAMATAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 4, "", Trim(rPajak![KD_KELURAHAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 5, "", Trim(rPajak![NM_KELURAHAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 6, "", Trim(rPajak![KD_KECAMATAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 7, "", Trim(rPajak![NM_KECAMATAN])
'            vBumi.ListItems.Item(I).ListSubItems.Add 8, "", Format(Trim(rPajak![LUAS_BUMI]), "#,#0")
'            vBumi.ListItems.Item(I).ListSubItems.Add 9, "", Format(Trim(rPajak![NILAI_SISTEM_BUMI]), "#,#0")
'            vBumi.ListItems.Item(I).ListSubItems.Add 10, "", Trim(rPajak![JNS_BUMI])
'            If Trim(rPajak![JNS_BUMI]) = 1 Then
'                JTANAH = "TANAH DAN BANGUNAN"
'            ElseIf Trim(rPajak![JNS_BUMI]) = 2 Then
'                JTANAH = "KAVLING DAN SIAP BANGUN"
'            ElseIf Trim(rPajak![JNS_BUMI]) = 3 Then
'                JTANAH = "TANAH KOSONG"
'            Else
'                JTANAH = "LAINNYA"
'            End If
'            vBumi.ListItems.Item(I).ListSubItems.Add 11, "", JTANAH
'            vBumi.ListItems.Item(I).ListSubItems.Add 12, "", Trim(rPajak![NO_FORMULIR])
'            vBumi.ListItems.Item(I).ListSubItems.Add 13, "", Trim(rPajak![KD_SEKTOR])
'            vBumi.ListItems.Item(I).ListSubItems.Add 14, "", Trim(rPajak![NM_SEKTOR])
'            vBumi.ListItems.Item(I).ListSubItems.Add 15, "", Trim(rPajak![KD_MAP])
'    rPajak.MoveNext
'    Loop
    '----------------------
    
    End If
    
     i = 0
    Do While Not rPajak.EOF
    'If rPajak!NM_KELURAHAN Like "*" & Trim(tCari.Text) & "*" Then
    i = i + 1
    vBumi.ListItems.Add i, "", Format(i, "#")
    vBumi.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
    'NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP)
            vBumi.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!NOPQ
            xxKel = Trim(rPajak![NM_KELURAHAN])
            If xxKel = True Or xxKel = "" Then xxKel = "-"
            xxKec = Trim(rPajak![NM_KECAMATAN])
            If xxKec = True Or xxKec = "" Then xxKec = "-"
            vBumi.ListItems.Item(i).ListSubItems.Add 3, "", xxKel & ", KEC. " & xxKec 'Trim(rPajak![NM_KELURAHAN]) & ", KEC. " & Trim(rPajak![NM_KECAMATAN])
            vBumi.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![KD_KELURAHAN])
            vBumi.ListItems.Item(i).ListSubItems.Add 5, "", Trim(rPajak![NM_KELURAHAN])
            vBumi.ListItems.Item(i).ListSubItems.Add 6, "", Trim(rPajak![KD_KECAMATAN])
            vBumi.ListItems.Item(i).ListSubItems.Add 7, "", Trim(rPajak![NM_KECAMATAN])
            xxLuas = Trim(rPajak![TOTAL_LUAS_BUMI])
            If IsNull(xxLuas) = True Or xxLuas = "" Then xxLuas = 0
            vBumi.ListItems.Item(i).ListSubItems.Add 8, "", Format(xxLuas, "#,#0") 'Format(Trim(rPajak![LUAS_BUMI]), "#,#0")
            xxNilai = Trim(rPajak![NILAI_SISTEM_BUMI])
            If IsNull(xxNilai) = True Or xxNilai = "" Then xxNilai = 0
            vBumi.ListItems.Item(i).ListSubItems.Add 9, "", Format(xxNilai, "#,#0") 'Format(Trim(rPajak![NILAI_SISTEM_BUMI]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 10, "", Trim(rPajak![JNS_BUMI])
            If Trim(rPajak![JNS_BUMI]) = 1 Then
                JTANAH = "TANAH DAN BANGUNAN"
            ElseIf Trim(rPajak![JNS_BUMI]) = 2 Then
                JTANAH = "KAVLING DAN SIAP BANGUN"
            ElseIf Trim(rPajak![JNS_BUMI]) = 3 Then
                JTANAH = "TANAH KOSONG"
            Else
                JTANAH = "FASILITAS UMUM"
            End If
            vBumi.ListItems.Item(i).ListSubItems.Add 11, "", JTANAH
            xxForm = Trim(rPajak![NO_FORMULIR_SPOP])
            If IsNull(xxForm) = True Or xxForm = "" Then xxForm = "-"
            vBumi.ListItems.Item(i).ListSubItems.Add 12, "", xxForm 'Trim(rPajak![NO_FORMULIR])
            xxID = Trim(rPajak![SUBJEK_PAJAK_ID])
            If xxID = "" Or IsNull(xxID) = True Then xxID = "-"
            vBumi.ListItems.Item(i).ListSubItems.Add 13, "", xxID 'Trim(rPajak![KD_SEKTOR])
            xxStatus = Trim(rPajak![KD_STATUS_WP])
            If xxStatus = "" Or IsNull(xxStatus) = True Then xxStatus = "0"
            vBumi.ListItems.Item(i).ListSubItems.Add 14, "", xxStatus 'xxNSektor 'Trim(rPajak![NM_SEKTOR])
            xxPersil = Trim(rPajak![NO_PERSIL])
            If xxPersil = "" Or IsNull(xxPersil) = True Then xxPersil = "0"
            vBumi.ListItems.Item(i).ListSubItems.Add 15, "", xxPersil 'Trim(rPajak![KD_MAP])
            xxZNT = Trim(rPajak![KD_ZNT])
            If xxZNT = "" Or IsNull(xxZNT) = True Then xxZNT = "00"
            vBumi.ListItems.Item(i).ListSubItems.Add 16, "", xxZNT 'Trim(rPajak![KD_MAP])
            xxJalan = Trim(rPajak![JALAN_OP])
            If xxJalan = "" Or IsNull(xxJalan) = True Then xxJalan = "-"
            vBumi.ListItems.Item(i).ListSubItems.Add 17, "", xxJalan 'Trim(rPajak![KD_MAP])
            xxBlok = Trim(rPajak![BLOK_KAV_NO_OP])
            If xxBlok = "" Or IsNull(xxBlok) = True Then xxBlok = "00"
            vBumi.ListItems.Item(i).ListSubItems.Add 18, "", xxBlok 'Trim(rPajak![KD_MAP])
            xxRW = Trim(rPajak![RW_OP])
            If xxRW = "" Or IsNull(xxRW) = True Then xxRW = "00"
            vBumi.ListItems.Item(i).ListSubItems.Add 19, "", xxRW 'Trim(rPajak![KD_MAP])
            xxRT = Trim(rPajak![RT_OP])
            If xxRT = "" Or IsNull(xxRT) = True Then xxRT = "00"
            vBumi.ListItems.Item(i).ListSubItems.Add 20, "", xxRT 'Trim(rPajak![KD_MAP])
            'xxData1 = Trim(rPajak![TGL_PENDATAAN_OP])
            'If xxJalan = "" Or IsNull(xxJalan) = True Then xxJalan = "-"
            vBumi.ListItems.Item(i).ListSubItems.Add 21, "", rPajak![TGL_PENDATAAN_OP]
            If IsNull(rPajak![NIP_PENDATA]) = True Then
                vBumi.ListItems.Item(i).ListSubItems.Add 22, "", "-"
            Else
                vBumi.ListItems.Item(i).ListSubItems.Add 22, "", rPajak![NIP_PENDATA]
            End If
            vBumi.ListItems.Item(i).ListSubItems.Add 23, "", rPajak![TGL_PEMERIKSAAN_OP]
            If IsNull(rPajak![NIP_PEMERIKSA_OP]) = True Then
                vBumi.ListItems.Item(i).ListSubItems.Add 24, "", "-"
            Else
                vBumi.ListItems.Item(i).ListSubItems.Add 24, "", rPajak![NIP_PEMERIKSA_OP]
            End If
            vBumi.ListItems.Item(i).ListSubItems.Add 25, "", rPajak![TGL_PEREKAMAN_OP]
            If IsNull(rPajak![NIP_PEREKAM_OP]) = True Then
                vBumi.ListItems.Item(i).ListSubItems.Add 26, "", "-"
            Else
                vBumi.ListItems.Item(i).ListSubItems.Add 26, "", rPajak![NIP_PEREKAM_OP]
            End If
            vBumi.ListItems.Item(i).ListSubItems.Add 27, "", rPajak![KD_KECAMATAN]
            vBumi.ListItems.Item(i).ListSubItems.Add 28, "", rPajak![KD_KELURAHAN]
            vBumi.ListItems.Item(i).ListSubItems.Add 29, "", rPajak![KD_BLOK]
            vBumi.ListItems.Item(i).ListSubItems.Add 30, "", rPajak![NO_URUT]
            vBumi.ListItems.Item(i).ListSubItems.Add 31, "", rPajak![KD_JNS_OP]
            
    'End If
    rPajak.MoveNext
    Loop
'     Set DataGrid1.DataSource = rPajak
LFIN.Caption = ""
LFIN.Caption = "Hasil Pencarian Berdasarkan " & oKlas(Pilih - 1).Caption & " : " & vBumi.ListItems.Count & " File"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExit_Click()
Unload Me
xID = 0
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
If xID = 1 Then
'    frmPBB.tBumi(0).Text = vBumi.SelectedItem.ListSubItems(12).Text
'    frmPBB.tBumi(1).Text = vBumi.SelectedItem.ListSubItems(2).Text
'    frmPBB.tBumi(3).Text = vBumi.SelectedItem.ListSubItems(6).Text & "-" & vBumi.SelectedItem.ListSubItems(7).Text
'    frmPBB.tBumi(4).Text = vBumi.SelectedItem.ListSubItems(4).Text & "-" & vBumi.SelectedItem.ListSubItems(5).Text
'    frmPBB.tBumi(12).Text = vBumi.SelectedItem.ListSubItems(10).Text & "-" & vBumi.SelectedItem.ListSubItems(11).Text
'    frmPBB.tBumi(13).Text = vBumi.SelectedItem.ListSubItems(8).Text
'    frmPBB.tBumi(15).Text = vBumi.SelectedItem.ListSubItems(9).Text
    'MsgBox vBumi.SelectedItem.ListSubItems(8).Text
    frmObjek_Pajak_Bm.cboNOP(0).Text = vBumi.SelectedItem.ListSubItems(6).Text & "-" & vBumi.SelectedItem.ListSubItems(7).Text 'Kecamatan
    frmObjek_Pajak_Bm.cboNOP(1).Text = vBumi.SelectedItem.ListSubItems(4).Text & "-" & vBumi.SelectedItem.ListSubItems(5).Text 'Kelurahan
    frmObjek_Pajak_Bm.cboNOP(2).Text = vBumi.SelectedItem.ListSubItems(29).Text
    frmObjek_Pajak_Bm.cboNOP(3).Text = vBumi.SelectedItem.ListSubItems(30).Text
    frmObjek_Pajak_Bm.cboNOP(4).Text = vBumi.SelectedItem.ListSubItems(31).Text
    frmObjek_Pajak_Bm.tBumi(0).Text = vBumi.SelectedItem.ListSubItems(2).Text 'NOP
    frmObjek_Pajak_Bm.mBUMI.Text = vBumi.SelectedItem.ListSubItems(2).Text 'NOP
    frmObjek_Pajak_Bm.cboJenis.Text = frmObjek_Pajak_Bm.cboJenis.List((vBumi.SelectedItem.ListSubItems(10).Text * 1) - 1) 'Jenis tanah
    frmObjek_Pajak_Bm.tBumi(1).Text = vBumi.SelectedItem.ListSubItems(8).Text 'Luas tanah
    frmObjek_Pajak_Bm.txtPajak(1).Text = vBumi.SelectedItem.ListSubItems(12).Text 'Formulir/Dokumen
    frmObjek_Pajak_Bm.tID1.Text = vBumi.SelectedItem.ListSubItems(13).Text 'ID Subjek Pajak
    frmObjek_Pajak_Bm.cboStatus.Text = frmObjek_Pajak_Bm.cboStatus.List((vBumi.SelectedItem.ListSubItems(14).Text * 1) - 1) 'Status Kepemilikan
    frmObjek_Pajak_Bm.tBumi(6).Text = vBumi.SelectedItem.ListSubItems(16).Text 'ZNT
    frmObjek_Pajak_Bm.cboJalan.Text = vBumi.SelectedItem.ListSubItems(16).Text & "-" & vBumi.SelectedItem.ListSubItems(17).Text 'Nama Jalan
    frmObjek_Pajak_Bm.tBumi(8).Text = vBumi.SelectedItem.ListSubItems(19).Text 'RW
    frmObjek_Pajak_Bm.tBumi(9).Text = vBumi.SelectedItem.ListSubItems(20).Text 'RT
    frmObjek_Pajak_Bm.tBumi(10).Text = vBumi.SelectedItem.ListSubItems(15).Text 'Persil
    frmObjek_Pajak_Bm.tBumi(11).Text = vBumi.SelectedItem.ListSubItems(18).Text 'Blok/Kav
    frmObjek_Pajak_Bm.dtPajak(0).Value = Format(vBumi.SelectedItem.ListSubItems(21).Text, "dd/mm/yyyy") 'Tanggal Pendataan
    frmObjek_Pajak_Bm.dtPajak(1).Value = Format(vBumi.SelectedItem.ListSubItems(23).Text, "dd/mm/yyyy") 'Tanggal Pemeriksaan
    frmObjek_Pajak_Bm.dtPajak(2).Value = Format(vBumi.SelectedItem.ListSubItems(25).Text, "dd/mm/yyyy") 'Tanggal Perekaman
    frmObjek_Pajak_Bm.tBumi(23).Text = vBumi.SelectedItem.ListSubItems(22).Text 'NIP Pendata
    frmObjek_Pajak_Bm.tBumi(24).Text = vBumi.SelectedItem.ListSubItems(24).Text 'NIP Pemeriksa
    frmObjek_Pajak_Bm.tBumi(25).Text = vBumi.SelectedItem.ListSubItems(26).Text 'NIP Perekam
    
    Unload Me

Else
    Unload Me
End If
xID = 0
End Sub

Private Sub Form_Activate()
On Error Resume Next
Screen.MousePointer = vbHourglass
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
Pilih = 1
oKlas(0).SetFocus

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
    tCari.Text = Rep(tCari.Text)
    KeyAscii = 0
    cmdCari_Click
   
End If


End Sub


Private Sub tCari_LostFocus()
tCari.Text = Rep(tCari.Text)
End Sub

Private Sub vBangunan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan.SortKey = ColumnHeader.Index - 1
vBangunan.Sorted = True
vBangunan.Sorted = False
vBangunan.SortOrder = lvwAscending

End Sub


Private Sub vBumi_Click()
''MsgBox vBumi.SelectedItem.ListSubItems(2).Text
'Screen.MousePointer = vbHourglass
'vBangunan.ListItems.Clear
'If vBumi.SelectedItem.ListSubItems(10).Text = 1 Then
''StringQ = "SELECT KD_PROPINSI, KD_DATI2, KD_KECAMATAN, KD_KELURAHAN, KD_BLOK, NO_URUT, KD_JNS_OP, NO_BNG, LUAS_BNG, NILAI_SISTEM_BNG, NO_FORMULIR_LSPOP, " & _
'vbCrLf & "KD_PROPINSI +'.'+ KD_DATI2 +'.'+ KD_KECAMATAN +'.'+ KD_KELURAHAN +'.'+ KD_BLOK + '-' + NO_URUT +'.'+ KD_JNS_OP AS NOPB FROM DAT_OP_BANGUNAN" '  WHERE NOPB='" & vBumi.SelectedItem.ListSubItems(2).Text & "'"
'StringQ = "Select * From vBANGUNAN WHERE NOPQ='" & vBumi.SelectedItem.ListSubItems(2).Text & "'"
'
'    openDB (StringQ)
'
'    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    I = 0
'    Do While Not rPajak.EOF
'    'If rPajak!NOPB = vBumi.SelectedItem.ListSubItems(2).Text Then
'    I = I + 1
'    vBangunan.ListItems.Add I, "", Format(I, "#")
'    vBangunan.ListItems.Item(I).ListSubItems.Add 1, "", Format(I, "#")
'    'NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP)
'            vBangunan.ListItems.Item(I).ListSubItems.Add 2, "", Trim(rPajak![kd_jpb])
'            vBangunan.ListItems.Item(I).ListSubItems.Add 3, "", rPajak!NM_JPB
'            vBangunan.ListItems.Item(I).ListSubItems.Add 4, "", Trim(rPajak![LUAS_BNG])
'            vBangunan.ListItems.Item(I).ListSubItems.Add 5, "", Trim(rPajak![NILAI_SISTEM_BNG])
'            vBangunan.ListItems.Item(I).ListSubItems.Add 6, "", Trim(rPajak![NO_BNG])
'            vBangunan.ListItems.Item(I).ListSubItems.Add 7, "", Trim(rPajak![NO_FORMULIR_LSPOP])
'
'    'End If
'    rPajak.MoveNext
'    Loop
'End If
'Screen.MousePointer = vbDefault
End Sub

Private Sub vBumi_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBumi.SortKey = ColumnHeader.Index - 1
vBumi.Sorted = True
vBumi.Sorted = False
vBumi.SortOrder = lvwAscending

End Sub

Private Sub oKlas_Click(Index As Integer)
On Error Resume Next
Screen.MousePointer = vbHourglass
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
Screen.MousePointer = vbDefault
End Sub
