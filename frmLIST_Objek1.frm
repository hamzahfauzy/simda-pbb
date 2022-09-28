VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLIST_Objek1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List : Subjek dan Objek Pajak"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11160
   Icon            =   "frmLIST_Objek1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   90
      Picture         =   "frmLIST_Objek1.frx":1CCA
      ScaleHeight     =   300
      ScaleWidth      =   10920
      TabIndex        =   15
      Top             =   1095
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
         Left            =   1590
         TabIndex        =   17
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
         Left            =   4125
         TabIndex        =   16
         Top             =   75
         Width           =   6765
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   75
      Picture         =   "frmLIST_Objek1.frx":6332
      ScaleHeight     =   300
      ScaleWidth      =   10905
      TabIndex        =   13
      Top             =   240
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
         TabIndex        =   14
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
      TabIndex        =   8
      Top             =   7170
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
      TabIndex        =   7
      Top             =   7170
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
      TabIndex        =   6
      Top             =   7170
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   4965
      TabIndex        =   11
      Top             =   465
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
         TabIndex        =   0
         Top             =   195
         Width           =   4305
      End
      Begin VB.CommandButton cmdCari 
         Height          =   375
         Left            =   5520
         Picture         =   "frmLIST_Objek1.frx":A99A
         Style           =   1  'Graphical
         TabIndex        =   1
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
         TabIndex        =   12
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
      TabIndex        =   9
      Top             =   465
      Width           =   4860
      Begin VB.OptionButton oKlas 
         Caption         =   "LOKASI OP"
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
         Index           =   3
         Left            =   3510
         Picture         =   "frmLIST_Objek1.frx":ACF5
         TabIndex        =   5
         Top             =   315
         Width           =   1320
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
         Index           =   0
         Left            =   120
         Picture         =   "frmLIST_Objek1.frx":191DE
         TabIndex        =   2
         Top             =   300
         Width           =   900
      End
      Begin VB.OptionButton oKlas 
         Caption         =   "NAMA WP"
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
         Left            =   1080
         Picture         =   "frmLIST_Objek1.frx":276C7
         TabIndex        =   3
         Top             =   300
         Width           =   1185
      End
      Begin VB.OptionButton oKlas 
         Caption         =   "LOKASI WP"
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
         Left            =   2250
         Picture         =   "frmLIST_Objek1.frx":35BB0
         TabIndex        =   4
         Top             =   315
         Width           =   1320
      End
   End
   Begin MSComctlLib.ListView vBumi 
      Height          =   5580
      Left            =   60
      TabIndex        =   10
      Top             =   1410
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   9843
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
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   22
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
         Key             =   "b"
         Text            =   "Nama Wajib Pajak"
         Object.Width           =   4234
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Alamat Wajib Pajak"
         Object.Width           =   6880
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Lokasi Objek Pajak"
         Object.Width           =   6880
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "L_BUMI"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "L_BNG"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "NJOP BUMI"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "NJOP BNG"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "TOTAL NJOP"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "NJOPTK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "SUBJEK_ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "JALAN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "KEL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "BLOK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "RW"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "RT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "KOTA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "POS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "NPWP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Persil"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   9195
      Left            =   -90
      Picture         =   "frmLIST_Objek1.frx":44099
      Stretch         =   -1  'True
      Top             =   -435
      Width           =   12960
   End
End
Attribute VB_Name = "frmLIST_Objek1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pilih
Dim xMIN(2), xMAX(2)
Dim xTKP(2)
Dim xBLOK, xRT, xRW, xJALAN, xKota, xLurah, xPos, xNPWP
Private Sub cmdCari_Click()
On Error GoTo Salah
'On Error Resume Next
CALL_NJOPTKP
vBumi.Sorted = False
    vBumi.Sorted = False
    Screen.MousePointer = vbHourglass
    
    Set rPajak = Nothing
    
    vBumi.ListItems.Clear
    
    If Pilih = 1 Then
    'StringQ = "SELECT DAT_OBJEK_PAJAK.KD_PROPINSI, DAT_OBJEK_PAJAK.KD_DATI2, DAT_OBJEK_PAJAK.KD_Kecamatan, DAT_OBJEK_PAJAK.KD_KELURAHAN, DAT_OBJEK_PAJAK.KD_BLOK, DAT_OBJEK_PAJAK.NO_URUT, DAT_OBJEK_PAJAK.KD_JNS_OP, DAT_OBJEK_PAJAK.SUBJEK_PAJAK_ID, DAT_OBJEK_PAJAK.NO_FORMULIR_SPOP, DAT_OBJEK_PAJAK.NO_PERSIL, " & _
"DAT_OBJEK_PAJAK.JALAN_OP, DAT_OBJEK_PAJAK.BLOK_KAV_NO_OP, DAT_OBJEK_PAJAK.RW_OP, DAT_OBJEK_PAJAK.RT_OP, DAT_OBJEK_PAJAK.TOTAL_LUAS_BUMI, DAT_OBJEK_PAJAK.TOTAL_LUAS_BNG, DAT_OBJEK_PAJAK.NJOP_BUMI, DAT_OBJEK_PAJAK.NJOP_BNG, DAT_OBJEK_PAJAK.JNS_TRANSAKSI_OP, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, " & _
"[DAT_OBJEK_PAJAK].[KD_PROPINSI]+'.'+[DAT_OBJEK_PAJAK].[KD_DATI2]+'.'+[DAT_OBJEK_PAJAK].[KD_KECAMATAN]+'.'+[DAT_OBJEK_PAJAK].[KD_KELURAHAN]+'.'+[DAT_OBJEK_PAJAK].[KD_BLOK]+'-'+[DAT_OBJEK_PAJAK].[NO_URUT]+'.'+[DAT_OBJEK_PAJAK].[KD_JNS_OP] AS NOPQ, DAT_SUBJEK_PAJAK.NM_WP, DAT_SUBJEK_PAJAK.JALAN_WP, DAT_SUBJEK_PAJAK.BLOK_KAV_NO_WP," & _
"DAT_SUBJEK_PAJAK.RW_WP, DAT_SUBJEK_PAJAK.RT_WP, DAT_SUBJEK_PAJAK.KELURAHAN_WP, DAT_SUBJEK_PAJAK.KOTA_WP , DAT_SUBJEK_PAJAK.NPWP FROM ((DAT_OBJEK_PAJAK INNER JOIN REF_KECAMATAN ON DAT_OBJEK_PAJAK.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) INNER JOIN REF_KELURAHAN ON (DAT_OBJEK_PAJAK.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (DAT_OBJEK_PAJAK.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN DAT_SUBJEK_PAJAK ON DAT_OBJEK_PAJAK.SUBJEK_PAJAK_ID = DAT_SUBJEK_PAJAK.SUBJEK_PAJAK_ID where REF_KELURAHAN.NM_KELURAHAN LIKE '" & "%" & Trim(tCari.Text) & "%" & "'"

    StringQ = "select * from QOBJEKPAJAK where NOPQ LIKE '" & "%" & Trim(tCari.Text) & "%" & "' AND (JNS_BUMI='1' OR JNS_BUMI='4')" ' AND TOTAL_LUAS_BNG>0)  "
    openDB (StringQ)
    
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
    i = i + 1
    vBumi.ListItems.Add i, "", Format(i, "#")
    vBumi.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
    'NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP)
    If IsNull(Trim(rPajak!BLOK_KAV_NO_WP)) = True Or rPajak!BLOK_KAV_NO_WP = "" Then
        xBLOK = "00"
    Else
        xBLOK = Trim(rPajak!BLOK_KAV_NO_WP)
    End If
    If IsNull(Trim(rPajak!RT_WP)) = True Or rPajak!RT_WP = "" Then
        xRT = "00"
    Else
        xRT = Trim(rPajak!RT_WP)
    End If
    If IsNull(Trim(rPajak!RW_WP)) = True Or rPajak!RW_WP = "" Then
        xRW = "00"
    Else
        xRW = Trim(rPajak!RW_WP)
    End If
    If IsNull(Trim(rPajak!JALAN_WP)) = True Or rPajak!JALAN_WP = "" Then
        xJALAN = "-"
    Else
        xJALAN = Trim(rPajak!JALAN_WP)
    End If
    If IsNull(Trim(rPajak!KOTA_WP)) = True Or rPajak!KOTA_WP = "" Then
        xKota = "-"
    Else
        xKota = Trim(rPajak!KOTA_WP)
    End If
    If IsNull(Trim(rPajak!KELURAHAN_WP)) = True Or rPajak!KELURAHAN_WP = "" Then
        xLurah = "-"
    Else
        xLurah = Trim(rPajak!KELURAHAN_WP)
    End If
    If IsNull(Trim(rPajak!KD_POS_WP)) = True Or rPajak!KD_POS_WP = "" Then
        xPos = "-"
    Else
        xPos = Trim(rPajak!KD_POS_WP)
    End If
    If IsNull(Trim(rPajak!NPWP)) = True Or rPajak!NPWP = "" Then
        xNPWP = "-"
    Else
        xNPWP = Trim(rPajak!NPWP)
    End If
    
    
    If IsNull(Trim(rPajak!BLOK_KAV_NO_OP)) = True Or rPajak!BLOK_KAV_NO_OP = "" Then
        xBlok1 = "-"
    Else
        xBlok1 = Trim(rPajak!BLOK_KAV_NO_OP)
    End If
    If IsNull(Trim(rPajak!RT_OP)) = True Or rPajak!RT_OP = "" Then
        xRT1 = "-"
    Else
        xRT1 = Trim(rPajak!RT_OP)
    End If
    If IsNull(Trim(rPajak!RW_OP)) = True Or rPajak!RW_OP = "" Then
        xRW1 = "-"
    Else
        xRW1 = Trim(rPajak!RW_OP)
    End If
    
    If IsNull(Trim(rPajak!NO_PERSIL)) = True Or rPajak!NO_PERSIL = "" Then
        xPersil = "-"
    Else
        xPersil = Trim(rPajak!NO_PERSIL)
    End If
    vBumi.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!NOPQ
    If IsNull(Trim(rPajak!Nm_wp)) = True Or rPajak!Nm_wp = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 3, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![Nm_wp])
    End If
     If IsNull(Trim(rPajak!JALAN_WP)) = True Or rPajak!JALAN_WP = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 4, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![JALAN_WP]) & ", " & Trim(rPajak![KELURAHAN_WP]) & ", BLOK: " & xBLOK & " /RW: " & xRW & "/RT: " & xRT & "-" & Trim(rPajak![KOTA_WP])
    End If
    If IsNull(Trim(rPajak!JALAN_OP)) = True Or rPajak!JALAN_OP = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 5, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 5, "", Trim(rPajak![JALAN_OP]) & ", " & Trim(rPajak![NM_KELURAHAN]) & ", BLOK: " & xBlok1 & " /RW: " & xRW1 & "/RT: " & xRT1 & " KEC. " & Trim(rPajak![NM_KECAMATAN])
    End If
            'vBumi.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![JALAN_WP]) & ", " & Trim(rPajak![KELURAHAN_WP]) & " BLOK " & Trim(rPajak![BLOK_KAV_NO_WP]) & " RW " & Trim(rPajak![RW_WP]) & "/RT " & Trim(rPajak![RT_WP]) & "-" & Trim(rPajak![KOTA_WP])
            
            
            vBumi.ListItems.Item(i).ListSubItems.Add 6, "", Format(Trim(rPajak![TOTAL_LUAS_BUMI]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 7, "", Format(Trim(rPajak![TOTAL_LUAS_BNG]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 8, "", Format(Trim(rPajak![NJOP_BUMI]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 9, "", Format(Trim(rPajak![NJOP_BNG]), "#,#0")
'            If Trim(rPajak![JNS_BUMI]) = 1 Then
'                JTANAH = "TANAH DAN BANGUNAN"
'            ElseIf Trim(rPajak![JNS_BUMI]) = 2 Then
'                JTANAH = "KAVLING DAN SIAP BANGUN"
'            ElseIf Trim(rPajak![JNS_BUMI]) = 3 Then
'                JTANAH = "TANAH KOSONG"
'            Else
'                JTANAH = "LAINNYA"
'            End If
            totnjop = Trim(rPajak![NJOP_BUMI]) * 1 + Trim(rPajak![NJOP_BNG]) * 1
            vBumi.ListItems.Item(i).ListSubItems.Add 10, "", Format(totnjop, "#,#0")
'            If totnjop < 500000000 Then
'                xNJOPTKP = 10000000
'            Else
'                xNJOPTKP = 15000000
'            End If
            
            If totnjop > xMIN(1) And totnjop <= xMAX(1) Then
                 xNJOPTKP = xTKP(1)
            Else
                 xNJOPTKP = xTKP(2)
            End If
            If rPajak![TOTAL_LUAS_BNG] <= 0 Or rPajak![NJOP_BNG] <= 0 Then
                xNJOPTKP = 0
            End If
'            xPBB = totNJOP - xNJOPTKP
'            If xPBB < 10000000 Then
'                xTarif = 3000
'            Else
'                xTarif = 0.003 * xPBB
'            End If
            vBumi.ListItems.Item(i).ListSubItems.Add 11, "", Format(xNJOPTKP, "#,#0")
    If IsNull(Trim(rPajak!SUBJEK_PAJAK_ID)) = True Or rPajak!SUBJEK_PAJAK_ID = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 12, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 12, "", Trim(rPajak!SUBJEK_PAJAK_ID)
    End If
            
            'vBumi.ListItems.Item(i).ListSubItems.Add 12, "", 0.003 * s
            vBumi.ListItems.Item(i).ListSubItems.Add 13, "", xJALAN
            vBumi.ListItems.Item(i).ListSubItems.Add 14, "", xLurah
            vBumi.ListItems.Item(i).ListSubItems.Add 15, "", xBLOK
            vBumi.ListItems.Item(i).ListSubItems.Add 16, "", xRW
            vBumi.ListItems.Item(i).ListSubItems.Add 17, "", xRT
            vBumi.ListItems.Item(i).ListSubItems.Add 18, "", xKota
            vBumi.ListItems.Item(i).ListSubItems.Add 19, "", xPos
            vBumi.ListItems.Item(i).ListSubItems.Add 20, "", xNPWP
            vBumi.ListItems.Item(i).ListSubItems.Add 21, "", xPersil
    rPajak.MoveNext
    Loop
    '----------------------
    ElseIf Pilih = 2 Then
   StringQ = "select * from QOBJEKPAJAK where NM_WP LIKE '" & "%" & Trim(tCari.Text) & "%" & "'"
    openDB (StringQ)
    
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
    i = i + 1
    vBumi.ListItems.Add i, "", Format(i, "#")
    vBumi.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
            vBumi.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!NOPQ
            vBumi.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![Nm_wp])
    If IsNull(Trim(rPajak!BLOK_KAV_NO_WP)) = True Or rPajak!BLOK_KAV_NO_WP = "" Then
        xBLOK = "00"
    Else
        xBLOK = Trim(rPajak!BLOK_KAV_NO_WP)
    End If
    If IsNull(Trim(rPajak!RT_WP)) = True Or rPajak!RT_WP = "" Then
        xRT = "00"
    Else
        xRT = Trim(rPajak!RT_WP)
    End If
    If IsNull(Trim(rPajak!RW_WP)) = True Or rPajak!RW_WP = "" Then
        xRW = "00"
    Else
        xRW = Trim(rPajak!RW_WP)
    End If
    If IsNull(Trim(rPajak!JALAN_WP)) = True Or rPajak!JALAN_WP = "" Then
        xJALAN = "-"
    Else
        xJALAN = Trim(rPajak!JALAN_WP)
    End If
    If IsNull(Trim(rPajak!KOTA_WP)) = True Or rPajak!KOTA_WP = "" Then
        xKota = "-"
    Else
        xKota = Trim(rPajak!KOTA_WP)
    End If
    If IsNull(Trim(rPajak!KELURAHAN_WP)) = True Or rPajak!KELURAHAN_WP = "" Then
        xLurah = "-"
    Else
        xLurah = Trim(rPajak!KELURAHAN_WP)
    End If
    If IsNull(Trim(rPajak!KD_POS_WP)) = True Or rPajak!KD_POS_WP = "" Then
        xPos = "-"
    Else
        xPos = Trim(rPajak!KD_POS_WP)
    End If
    If IsNull(Trim(rPajak!NPWP)) = True Or rPajak!NPWP = "" Then
        xNPWP = "-"
    Else
        xNPWP = Trim(rPajak!NPWP)
    End If
    
    
    
    If IsNull(Trim(rPajak!BLOK_KAV_NO_OP)) = True Or rPajak!BLOK_KAV_NO_OP = "" Then
        xBlok1 = "-"
    Else
        xBlok1 = Trim(rPajak!BLOK_KAV_NO_OP)
    End If
    If IsNull(Trim(rPajak!RT_OP)) = True Or rPajak!RT_OP = "" Then
        xRT1 = "-"
    Else
        xRT1 = Trim(rPajak!RT_OP)
    End If
    If IsNull(Trim(rPajak!RW_OP)) = True Or rPajak!RW_OP = "" Then
        xRW1 = "-"
    Else
        xRW1 = Trim(rPajak!RW_OP)
    End If
    If IsNull(Trim(rPajak!NO_PERSIL)) = True Or rPajak!NO_PERSIL = "" Then
        xPersil = "-"
    Else
        xPersil = Trim(rPajak!NO_PERSIL)
    End If
    
    vBumi.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!NOPQ
            
            If IsNull(Trim(rPajak!Nm_wp)) = True Or rPajak!Nm_wp = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 3, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![Nm_wp])
    End If
     If IsNull(Trim(rPajak!JALAN_WP)) = True Or rPajak!JALAN_WP = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 4, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![JALAN_WP]) & ", " & Trim(rPajak![KELURAHAN_WP]) & ", BLOK: " & xBLOK & " /RW: " & xRW & "/RT: " & xRT & "-" & Trim(rPajak![KOTA_WP])
    End If
    If IsNull(Trim(rPajak!JALAN_OP)) = True Or rPajak!JALAN_OP = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 5, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 5, "", Trim(rPajak![JALAN_OP]) & ", " & Trim(rPajak![NM_KELURAHAN]) & ", BLOK: " & xBlok1 & " /RW: " & xRW1 & "/RT: " & xRT1 & " KEC. " & Trim(rPajak![NM_KECAMATAN])
    End If
    
            vBumi.ListItems.Item(i).ListSubItems.Add 6, "", Format(Trim(rPajak![TOTAL_LUAS_BUMI]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 7, "", Format(Trim(rPajak![TOTAL_LUAS_BNG]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 8, "", Format(Trim(rPajak![NJOP_BUMI]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 9, "", Format(Trim(rPajak![NJOP_BNG]), "#,#0")

            totnjop = Trim(rPajak![NJOP_BUMI]) * 1 + Trim(rPajak![NJOP_BNG]) * 1
            vBumi.ListItems.Item(i).ListSubItems.Add 10, "", Format(totnjop, "#,#0")
            
            If totnjop > xMIN(1) And totnjop <= xMAX(1) Then
                 xNJOPTKP = xTKP(1)
            Else
                 xNJOPTKP = xTKP(2)
            End If
            If rPajak![TOTAL_LUAS_BNG] <= 0 Or rPajak![NJOP_BNG] <= 0 Then
                xNJOPTKP = 0
            End If
            vBumi.ListItems.Item(i).ListSubItems.Add 11, "", Format(xNJOPTKP, "#,#0")
    If IsNull(Trim(rPajak!SUBJEK_PAJAK_ID)) = True Or rPajak!SUBJEK_PAJAK_ID = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 12, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 12, "", Trim(rPajak!SUBJEK_PAJAK_ID)
    End If
            vBumi.ListItems.Item(i).ListSubItems.Add 13, "", xJALAN
            vBumi.ListItems.Item(i).ListSubItems.Add 14, "", xLurah
            vBumi.ListItems.Item(i).ListSubItems.Add 15, "", xBLOK
            vBumi.ListItems.Item(i).ListSubItems.Add 16, "", xRW
            vBumi.ListItems.Item(i).ListSubItems.Add 17, "", xRT
            vBumi.ListItems.Item(i).ListSubItems.Add 18, "", xKota
            vBumi.ListItems.Item(i).ListSubItems.Add 19, "", xPos
            vBumi.ListItems.Item(i).ListSubItems.Add 20, "", xNPWP
            vBumi.ListItems.Item(i).ListSubItems.Add 21, "", xPersil
    rPajak.MoveNext
    Loop
    '-------------------
    ElseIf Pilih = 3 Then
    StringQ = "select * from QOBJEKPAJAK where [JALAN_WP] + ', ' + [KELURAHAN_WP] + ' BLOK ' + [BLOK_KAV_NO_WP] + ' RW ' + [RW_WP] + '/RT ' + [RT_WP] + '-' + [KOTA_WP] LIKE '" & "%" & Trim(tCari.Text) & "%" & "'"
    openDB (StringQ)
    
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
    i = i + 1
    vBumi.ListItems.Add i, "", Format(i, "#")
    vBumi.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
            vBumi.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!NOPQ
            vBumi.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![Nm_wp])
    
    If IsNull(Trim(rPajak!BLOK_KAV_NO_WP)) = True Or rPajak!BLOK_KAV_NO_WP = "" Then
        xBLOK = "-"
    Else
        xBLOK = Trim(rPajak!BLOK_KAV_NO_WP)
    End If
    If IsNull(Trim(rPajak!RT_WP)) = True Or rPajak!RT_WP = "" Then
        xRT = "-"
    Else
        xRT = Trim(rPajak!RT_WP)
    End If
    If IsNull(Trim(rPajak!RW_WP)) = True Or rPajak!RW_WP = "" Then
        xRW = "-"
    Else
        xRW = Trim(rPajak!RW_WP)
    End If
    If IsNull(Trim(rPajak!JALAN_WP)) = True Or rPajak!JALAN_WP = "" Then
        xJALAN = "-"
    Else
        xJALAN = Trim(rPajak!JALAN_WP)
    End If
    If IsNull(Trim(rPajak!KOTA_WP)) = True Or rPajak!KOTA_WP = "" Then
        xKota = "-"
    Else
        xKota = Trim(rPajak!KOTA_WP)
    End If
    If IsNull(Trim(rPajak!KELURAHAN_WP)) = True Or rPajak!KELURAHAN_WP = "" Then
        xLurah = "-"
    Else
        xLurah = Trim(rPajak!KELURAHAN_WP)
    End If
    If IsNull(Trim(rPajak!KD_POS_WP)) = True Or rPajak!KD_POS_WP = "" Then
        xPos = "-"
    Else
        xPos = Trim(rPajak!KD_POS_WP)
    End If
    If IsNull(Trim(rPajak!NPWP)) = True Or rPajak!NPWP = "" Then
        xNPWP = "-"
    Else
        xNPWP = Trim(rPajak!NPWP)
    End If
    
    If IsNull(Trim(rPajak!BLOK_KAV_NO_OP)) = True Or rPajak!BLOK_KAV_NO_OP = "" Then
        xBlok1 = "-"
    Else
        xBlok1 = Trim(rPajak!BLOK_KAV_NO_OP)
    End If
    If IsNull(Trim(rPajak!RT_OP)) = True Or rPajak!RT_OP = "" Then
        xRT1 = "-"
    Else
        xRT1 = Trim(rPajak!RT_OP)
    End If
    If IsNull(Trim(rPajak!RW_OP)) = True Or rPajak!RW_OP = "" Then
        xRW1 = "-"
    Else
        xRW1 = Trim(rPajak!RW_OP)
    End If
    
     If IsNull(Trim(rPajak!JALAN_OP)) = True Or rPajak!JALAN_OP = "" Then
        xjalan1 = "-"
    Else
        xjalan1 = Trim(rPajak!JALAN_OP)
    End If
    
    If IsNull(Trim(rPajak!NM_KELURAHAN)) = True Or rPajak!NM_KELURAHAN = "" Then
        xLurah1 = "-"
    Else
        xLurah1 = Trim(rPajak!NM_KELURAHAN)
    End If
    If IsNull(Trim(rPajak!NM_KECAMATAN)) = True Or rPajak!NM_KECAMATAN = "" Then
        xCamat1 = "-"
    Else
        xCamat1 = Trim(rPajak!NM_KECAMATAN)
    End If
    If IsNull(Trim(rPajak!NO_PERSIL)) = True Or rPajak!NO_PERSIL = "" Then
        xPersil = "-"
    Else
        xPersil = Trim(rPajak!NO_PERSIL)
    End If
    
    'vBumi.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!NOPQ
     '       vBumi.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![Nm_wp])
            'vBumi.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![JALAN_WP]) & ", " & Trim(rPajak![KELURAHAN_WP]) & " BLOK " & Trim(rPajak![BLOK_KAV_NO_WP]) & " RW " & Trim(rPajak![RW_WP]) & "/RT " & Trim(rPajak![RT_WP]) & "-" & Trim(rPajak![KOTA_WP])
            vBumi.ListItems.Item(i).ListSubItems.Add 4, "", xJALAN & ", " & xLurah & ", BLOK: " & xBLOK & " /RW: " & xRW & "/RT: " & xRT & "-" & xKota 'Trim(rPajak![KOTA_WP])
            vBumi.ListItems.Item(i).ListSubItems.Add 5, "", xjalan1 & ", " & xLurah1 & ", BLOK: " & xBlok1 & " /RW: " & xRW1 & "/RT: " & xRT1 & " KEC. " & xCamat1 'Trim(rPajak![NM_KECAMATAN])
            
            vBumi.ListItems.Item(i).ListSubItems.Add 6, "", Format(Trim(rPajak![TOTAL_LUAS_BUMI]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 7, "", Format(Trim(rPajak![TOTAL_LUAS_BNG]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 8, "", Format(Trim(rPajak![NJOP_BUMI]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 9, "", Format(Trim(rPajak![NJOP_BNG]), "#,#0")

            totnjop = Trim(rPajak![NJOP_BUMI]) * 1 + Trim(rPajak![NJOP_BNG]) * 1
            vBumi.ListItems.Item(i).ListSubItems.Add 10, "", Format(totnjop, "#,#0")
            
            If totnjop > xMIN(1) And totnjop <= xMAX(1) Then
                 xNJOPTKP = xTKP(1)
            Else
                 xNJOPTKP = xTKP(2)
            End If
            If rPajak![TOTAL_LUAS_BNG] <= 0 Or rPajak![NJOP_BNG] <= 0 Then
                xNJOPTKP = 0
            End If
            vBumi.ListItems.Item(i).ListSubItems.Add 11, "", Format(xNJOPTKP, "#,#0")
            If IsNull(Trim(rPajak!SUBJEK_PAJAK_ID)) = True Or rPajak!SUBJEK_PAJAK_ID = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 12, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 12, "", Trim(rPajak!SUBJEK_PAJAK_ID)
    End If
            vBumi.ListItems.Item(i).ListSubItems.Add 13, "", xJALAN
            vBumi.ListItems.Item(i).ListSubItems.Add 14, "", xLurah
            vBumi.ListItems.Item(i).ListSubItems.Add 15, "", xBLOK
            vBumi.ListItems.Item(i).ListSubItems.Add 16, "", xRW
            vBumi.ListItems.Item(i).ListSubItems.Add 17, "", xRT
            vBumi.ListItems.Item(i).ListSubItems.Add 18, "", xKota
            vBumi.ListItems.Item(i).ListSubItems.Add 19, "", xPos
            vBumi.ListItems.Item(i).ListSubItems.Add 20, "", xNPWP
            vBumi.ListItems.Item(i).ListSubItems.Add 21, "", xPersil
            rPajak.MoveNext
    Loop
    '----------------------
    ElseIf Pilih = 4 Then
    'StringQ = "select * from QOBJEKPAJAK where [JALAN_OP] + ', ' + [NM_KELURAHAN] + ' BLOK ' + [BLOK_KAV_NO_OP] + ' RW ' + [RW_OP] + '/RT ' + [RT_OP] + '-' + [NM_KECAMATAN] LIKE '" & "%" & Trim(tCari.Text) & "%" & "'"
    StringQ = "select * from QOBJEKPAJAK where [JALAN_OP] + ', ' + [NM_KELURAHAN] + ' BLOK ' + [BLOK_KAV_NO_OP] + ' RW ' + [RW_OP] + '/RT ' + [RT_OP] + '-' + [NM_KECAMATAN] LIKE '" & "%" & Trim(tCari.Text) & "%" & "'"
    openDB (StringQ)
    
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
    i = i + 1
    
    vBumi.ListItems.Add i, "", Format(i, "#")
    vBumi.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
    If IsNull(Trim(rPajak!BLOK_KAV_NO_WP)) = True Or rPajak!BLOK_KAV_NO_WP = "" Then
        xBLOK = "00"
    Else
        xBLOK = Trim(rPajak!BLOK_KAV_NO_WP)
    End If
    If IsNull(Trim(rPajak!RT_WP)) = True Or rPajak!RT_WP = "" Then
        xRT = "00"
    Else
        xRT = Trim(rPajak!RT_WP)
    End If
    If IsNull(Trim(rPajak!RW_WP)) = True Or rPajak!RW_WP = "" Then
        xRW = "00"
    Else
        xRW = Trim(rPajak!RW_WP)
    End If
    If IsNull(Trim(rPajak!JALAN_WP)) = True Or rPajak!JALAN_WP = "" Then
        xJALAN = "-"
    Else
        xJALAN = Trim(rPajak!JALAN_WP)
    End If
    If IsNull(Trim(rPajak!KOTA_WP)) = True Or rPajak!KOTA_WP = "" Then
        xKota = "-"
    Else
        xKota = Trim(rPajak!KOTA_WP)
    End If
    If IsNull(Trim(rPajak!KELURAHAN_WP)) = True Or rPajak!KELURAHAN_WP = "" Then
        xLurah = "-"
    Else
        xLurah = Trim(rPajak!KELURAHAN_WP)
    End If
    If IsNull(Trim(rPajak!KD_POS_WP)) = True Or rPajak!KD_POS_WP = "" Then
        xPos = "-"
    Else
        xPos = Trim(rPajak!KD_POS_WP)
    End If
    If IsNull(Trim(rPajak!NPWP)) = True Or rPajak!NPWP = "" Then
        xNPWP = "-"
    Else
        xNPWP = Trim(rPajak!NPWP)
    End If
    
    If IsNull(Trim(rPajak!BLOK_KAV_NO_OP)) = True Or rPajak!BLOK_KAV_NO_OP = "" Then
        xBlok1 = "-"
    Else
        xBlok1 = Trim(rPajak!BLOK_KAV_NO_OP)
    End If
    If IsNull(Trim(rPajak!RT_OP)) = True Or rPajak!RT_OP = "" Then
        xRT1 = "-"
    Else
        xRT1 = Trim(rPajak!RT_OP)
    End If
    If IsNull(Trim(rPajak!RW_OP)) = True Or rPajak!RW_OP = "" Then
        xRW1 = "-"
    Else
        xRW1 = Trim(rPajak!RW_OP)
    End If
    
     If IsNull(Trim(rPajak!JALAN_OP)) = True Or rPajak!JALAN_OP = "" Then
        xjalan1 = "-"
    Else
        xjalan1 = Trim(rPajak!JALAN_OP)
    End If
    
    If IsNull(Trim(rPajak!NM_KELURAHAN)) = True Or rPajak!NM_KELURAHAN = "" Then
        xLurah1 = "-"
    Else
        xLurah1 = Trim(rPajak!NM_KELURAHAN)
    End If
    If IsNull(Trim(rPajak!NM_KECAMATAN)) = True Or rPajak!NM_KECAMATAN = "" Then
        xCamat1 = "-"
    Else
        xCamat1 = Trim(rPajak!NM_KECAMATAN)
    End If
    If IsNull(Trim(rPajak!NO_PERSIL)) = True Or rPajak!NO_PERSIL = "" Then
        xPersil = "-"
    Else
        xPersil = Trim(rPajak!NO_PERSIL)
    End If
    
    vBumi.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!NOPQ
            If IsNull(Trim(rPajak!Nm_wp)) = True Or rPajak!Nm_wp = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 3, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![Nm_wp])
    End If
     
            'vBumi.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![JALAN_WP]) & ", " & Trim(rPajak![KELURAHAN_WP]) & " BLOK " & Trim(rPajak![BLOK_KAV_NO_WP]) & " RW " & Trim(rPajak![RW_WP]) & "/RT " & Trim(rPajak![RT_WP]) & "-" & Trim(rPajak![KOTA_WP])
            vBumi.ListItems.Item(i).ListSubItems.Add 4, "", xJALAN & ", " & xLurah & ", BLOK: " & xBLOK & " /RW: " & xRW & "/RT: " & xRT & "-" & xKota 'Trim(rPajak![KOTA_WP])
            vBumi.ListItems.Item(i).ListSubItems.Add 5, "", xjalan1 & ", " & xLurah1 & ", BLOK: " & xBlok1 & " /RW: " & xRW1 & "/RT: " & xRT1 & " KEC. " & xCamat1 'Trim(rPajak![NM_KECAMATAN])
            
'            vBumi.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!NOPQ
'            vBumi.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![Nm_wp])
'
'            vBumi.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![JALAN_WP]) & ", " & Trim(rPajak![Kelurahan_WP]) & " BLOK " & Trim(rPajak![BLOK_KAV_NO_WP]) & " RW " & Trim(rPajak![RW_WP]) & "/RT " & Trim(rPajak![RT_WP]) & "-" & Trim(rPajak![KOTA_WP])
'            vBumi.ListItems.Item(i).ListSubItems.Add 5, "", Trim(rPajak![JALAN_oP]) & ", " & Trim(rPajak![NM_Kelurahan]) & " BLOK " & Trim(rPajak![BLOK_KAV_NO_OP]) & " RW " & Trim(rPajak![RW_OP]) & "/RT " & Trim(rPajak![RT_OP]) & " KEC. " & Trim(rPajak![NM_KECAMATAN])
            vBumi.ListItems.Item(i).ListSubItems.Add 6, "", Format(Trim(rPajak![TOTAL_LUAS_BUMI]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 7, "", Format(Trim(rPajak![TOTAL_LUAS_BNG]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 8, "", Format(Trim(rPajak![NJOP_BUMI]), "#,#0")
            vBumi.ListItems.Item(i).ListSubItems.Add 9, "", Format(Trim(rPajak![NJOP_BNG]), "#,#0")

            totnjop = Trim(rPajak![NJOP_BUMI]) * 1 + Trim(rPajak![NJOP_BNG]) * 1
            vBumi.ListItems.Item(i).ListSubItems.Add 10, "", Format(totnjop, "#,#0")
            
            If totnjop > xMIN(1) And totnjop <= xMAX(1) Then
                 xNJOPTKP = xTKP(1)
            Else
                 xNJOPTKP = xTKP(2)
            End If
            If rPajak![TOTAL_LUAS_BNG] <= 0 Or rPajak![NJOP_BNG] <= 0 Then
                xNJOPTKP = 0
            End If
            vBumi.ListItems.Item(i).ListSubItems.Add 11, "", Format(xNJOPTKP, "#,#0")
            If IsNull(Trim(rPajak!SUBJEK_PAJAK_ID)) = True Or rPajak!SUBJEK_PAJAK_ID = "" Then
        vBumi.ListItems.Item(i).ListSubItems.Add 12, "", "-"
    Else
        vBumi.ListItems.Item(i).ListSubItems.Add 12, "", Trim(rPajak!SUBJEK_PAJAK_ID)
    End If
            vBumi.ListItems.Item(i).ListSubItems.Add 13, "", xJALAN
            vBumi.ListItems.Item(i).ListSubItems.Add 14, "", xLurah
            vBumi.ListItems.Item(i).ListSubItems.Add 15, "", xBLOK
            vBumi.ListItems.Item(i).ListSubItems.Add 16, "", xRW
            vBumi.ListItems.Item(i).ListSubItems.Add 17, "", xRT
            vBumi.ListItems.Item(i).ListSubItems.Add 18, "", xKota
            vBumi.ListItems.Item(i).ListSubItems.Add 19, "", xPos
            vBumi.ListItems.Item(i).ListSubItems.Add 20, "", xNPWP
            vBumi.ListItems.Item(i).ListSubItems.Add 21, "", xPersil

    rPajak.MoveNext
    Loop
    End If
LFIN.Caption = ""
LFIN.Caption = "Jumlah " & oKlas(Pilih - 1).Caption & " : " & vBumi.ListItems.Count & " File"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
On Error Resume Next
tCari.Text = ""
vBumi.ListItems.Clear
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
Unload Me
xID = 0
End Sub

Private Sub cmdOK_Click()
On Error GoTo Salah
If xID = 3 Then
'    frmPBB.tBumi(0).Text = vBumi.SelectedItem.ListSubItems(11).Text 'NOP
'    frmPBB.tBumi(1).Text = vBumi.SelectedItem.ListSubItems(2).Text 'Nama
'    frmPBB.tBumi(3).Text = vBumi.SelectedItem.ListSubItems(5).Text & "-" & vBumi.SelectedItem.ListSubItems(6).Text 'Kecamatan
'    frmPBB.tBumi(4).Text = vBumi.SelectedItem.ListSubItems(3).Text & "-" & vBumi.SelectedItem.ListSubItems(4).Text 'Kelurahan
'    frmPBB.tBumi(12).Text = vBumi.SelectedItem.ListSubItems(9).Text & "-" & vBumi.SelectedItem.ListSubItems(10).Text 'Lokasi
'    frmPBB.tBumi(13).Text = vBumi.SelectedItem.ListSubItems(7).Text
'    frmPBB.tBumi(15).Text = vBumi.SelectedItem.ListSubItems(8).Text
    
    frmObjek_Pajak_Bg.txtPajak(1).Text = vBumi.SelectedItem.ListSubItems(2).Text 'NOP
    frmObjek_Pajak_Bg.aNOP.Text = vBumi.SelectedItem.ListSubItems(2).Text 'NOP
    frmObjek_Pajak_Bg.txtPajak(2).Text = vBumi.SelectedItem.ListSubItems(12).Text '" NAMA" & vbTab & ": " & vBumi.SelectedItem.ListSubItems(3).Text & vbCrLf & " ALAMAT" & vbTab & ": " & vBumi.SelectedItem.ListSubItems(5).Text 'NAMA dan Alamat
    'frmObjek_Pajak_Bg.txtPajak(0).Text = vBumi.SelectedItem.ListSubItems(2).Text 'NOP
    frmObjek_Pajak_Bg.aNOP.SetFocus
    '& "-" & vBumi.SelectedItem.ListSubItems(6).Text 'Kecamatan
    Unload Me
ElseIf xID = 4 Then
    frmSPPT_Tunggal.tNOP(0).Text = vBumi.SelectedItem.ListSubItems(2).Text 'NOP
    frmSPPT_Tunggal.aNOP.Text = vBumi.SelectedItem.ListSubItems(2).Text
'    frmSPPT_Tunggal.tNOP(1).Text = vBumi.SelectedItem.ListSubItems(3).Text
'    frmSPPT_Tunggal.tNOP(11).Text = vBumi.SelectedItem.ListSubItems(6).Text
'    frmSPPT_Tunggal.tNOP(12).Text = vBumi.SelectedItem.ListSubItems(7).Text
'    frmSPPT_Tunggal.tNOP(15).Text = vBumi.SelectedItem.ListSubItems(8).Text
'    frmSPPT_Tunggal.tNOP(16).Text = vBumi.SelectedItem.ListSubItems(9).Text
'    frmSPPT_Tunggal.tNOP(17).Text = vBumi.SelectedItem.ListSubItems(10).Text
'    frmSPPT_Tunggal.tNOP(18).Text = vBumi.SelectedItem.ListSubItems(11).Text
'    frmSPPT_Tunggal.tNOP(2).Text = vBumi.SelectedItem.ListSubItems(13).Text
'    frmSPPT_Tunggal.tNOP(6).Text = vBumi.SelectedItem.ListSubItems(14).Text
'    frmSPPT_Tunggal.tNOP(3).Text = vBumi.SelectedItem.ListSubItems(15).Text
'    frmSPPT_Tunggal.tNOP(4).Text = vBumi.SelectedItem.ListSubItems(16).Text
'    frmSPPT_Tunggal.tNOP(5).Text = vBumi.SelectedItem.ListSubItems(17).Text
'    frmSPPT_Tunggal.tNOP(7).Text = vBumi.SelectedItem.ListSubItems(18).Text
'    frmSPPT_Tunggal.tNOP(8).Text = vBumi.SelectedItem.ListSubItems(19).Text
'    frmSPPT_Tunggal.tNOP(9).Text = vBumi.SelectedItem.ListSubItems(20).Text
'    frmSPPT_Tunggal.tID.Text = vBumi.SelectedItem.ListSubItems(12).Text
'    frmSPPT_Tunggal.tNOP(10).Text = vBumi.SelectedItem.ListSubItems(21).Text
    'frmSPPT_Tunggal.tNOP(1).Text = vBumi.SelectedItem.ListSubItems(12).Text '" NAMA" & vbTab & ": " & vBumi.SelectedItem.ListSubItems(3).Text & vbCrLf & " ALAMAT" & vbTab & ": " & vBumi.SelectedItem.ListSubItems(5).Text 'NAMA dan Alamat
    frmSPPT_Tunggal.aNOP.SetFocus
    frmSPPT_Tunggal.Show
    Unload Me
ElseIf xID = 5 Then
    frmBayar1.aNOP.Text = vBumi.SelectedItem.ListSubItems(2).Text 'NOP
    frmBayar1.aNOP.SetFocus
    frmBayar1.Show
    Unload Me
ElseIf xID = 6 Then
    frmCetak_Massal.aNOP.Text = vBumi.SelectedItem.ListSubItems(2).Text 'NOP
    frmCetak_Massal.aNOP.SetFocus
    frmCetak_Massal.Show
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
Pilih = 1
oKlas(1).SetFocus

tCari.SetFocus

Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
xID = 0
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





'Private Sub vBumi_Click()
''MsgBox vBumi.SelectedItem.ListSubItems(2).Text
'Screen.MousePointer = vbHourglass
'vBangunan.ListItems.Clear
'StringQ = "SELECT KD_PROPINSI, KD_DATI2, KD_KECAMATAN, KD_KELURAHAN, KD_BLOK, NO_URUT, KD_JNS_OP, NO_BNG, LUAS_BNG, NILAI_SISTEM_BNG, NO_FORMULIR_LSPOP, " & _
'vbCrLf & "KD_PROPINSI +'.'+ KD_DATI2 +'.'+ KD_KECAMATAN +'.'+ KD_KELURAHAN +'.'+ KD_BLOK + '-' + NO_URUT +'.'+ KD_JNS_OP AS NOPB FROM DAT_OP_BANGUNAN" '  WHERE NOPB='" & vBumi.SelectedItem.ListSubItems(2).Text & "'"
'
'
'    openDB (StringQ)
'
'    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    i = 0
'    Do While Not rPajak.EOF
'    If rPajak!NOPB = vBumi.SelectedItem.ListSubItems(2).Text Then
'    i = i + 1
'    vBangunan.ListItems.Add i, "", Format(i, "#")
'    vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
'    'NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP)
'            vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!NOPB
'            vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![LUAS_BNG])
'            vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![NILAI_SISTEM_BNG])
'            vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Trim(rPajak![NO_BNG])
'            vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", Trim(rPajak![NO_FORMULIR_LSPOP])
'    End If
'    rPajak.MoveNext
'    Loop
'Screen.MousePointer = vbDefault
'End Sub


Private Sub tCari_LostFocus()
tCari.Text = Rep(tCari.Text)
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

Sub CALL_NJOPTKP()
'On Error GoTo Salah
On Error Resume Next
xxSTR = "Select * From Tarif order by NJOP_MIN"
openDB (xxSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
Do While Not rPajak.EOF
    i = i + 1
    xMIN(i) = rPajak!NJOP_MIN
    xMAX(i) = rPajak!NJOP_MAX
    xTKP(i) = rPajak!NJOPTKP
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

