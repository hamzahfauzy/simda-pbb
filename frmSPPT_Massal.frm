VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSPPT_Massal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penetapan SPPT Massal"
   ClientHeight    =   4065
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Icon            =   "frmSPPT_Massal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5895
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      TabIndex        =   26
      Top             =   2895
      Width           =   5880
      Begin VB.ComboBox ccBayar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   8
         Top             =   45
         Width           =   4320
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Bayar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   27
         Top             =   75
         Width           =   1305
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   -30
      ScaleHeight     =   480
      ScaleWidth      =   5955
      TabIndex        =   21
      Top             =   0
      Width           =   5955
      Begin MSComctlLib.ProgressBar pNilai 
         Height          =   255
         Left            =   30
         TabIndex        =   24
         Top             =   15
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label LNilai 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "NOP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   25
         Top             =   270
         Width           =   5715
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
      Left            =   3240
      TabIndex        =   11
      Top             =   3615
      Width           =   915
   End
   Begin VB.CommandButton cmdCear 
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
      Left            =   2340
      TabIndex        =   10
      Top             =   3615
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Proses"
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
      Left            =   1440
      TabIndex        =   9
      Top             =   3615
      Width           =   915
   End
   Begin MSComctlLib.ListView vOP 
      Height          =   3960
      Left            =   6045
      TabIndex        =   23
      Top             =   0
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   6985
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
      MousePointer    =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   51
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "No"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "PROP"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "KAB"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "KEC"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "KEL"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "BLOK"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "URUT"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "JNS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "TPAJAK"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "SIKLUS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "KANWIL"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "KPBB"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "BANK1"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "BANK2"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "KD_TP"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "NAMA WP"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "ALAMAT WP"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   18
         Text            =   "KAV"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   19
         Text            =   "RW"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   20
         Text            =   "RT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "KELURAHAN"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   22
         Text            =   "KOTA"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Text            =   "POS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "NPWP"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   25
         Text            =   "NOPERSIL"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   26
         Text            =   "K_TNH"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   27
         Text            =   "THN"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   28
         Text            =   "K_BNG"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   29
         Text            =   "THN"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "J_TEMPO"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Text            =   "L_BUMI"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Text            =   "L_BNG"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   33
         Text            =   "NJOP_BM"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "NJOP_BNG"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   35
         Text            =   "TOTAL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   36
         Text            =   "NJOPTKP"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   37
         Text            =   "NJKP"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   38
         Text            =   "BAYAR"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   39
         Text            =   "KURANG"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   40
         Text            =   "TOTAL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "ST_BYR"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   42
         Text            =   "STS_TAGIH"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "CETAK"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "T_TERBIT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "T_CETAK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "NIP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "PROSES"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "FLAG_NJOPTKP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "NOP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "J_Bumi"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   -105
      ScaleHeight     =   840
      ScaleWidth      =   6150
      TabIndex        =   22
      Top             =   3330
      Width           =   6150
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   -30
      TabIndex        =   12
      Top             =   345
      Width           =   5985
      Begin VB.TextBox tTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1515
         TabIndex        =   7
         Text            =   "0"
         Top             =   2190
         Width           =   4260
      End
      Begin VB.TextBox tKurang 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3735
         TabIndex        =   6
         Text            =   "0"
         Top             =   1815
         Width           =   2040
      End
      Begin VB.ComboBox ccKel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1515
         TabIndex        =   2
         Top             =   855
         Width           =   4260
      End
      Begin VB.TextBox tSPPT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   1815
         Width           =   1290
      End
      Begin VB.ComboBox ccTahun 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1515
         TabIndex        =   0
         Top             =   195
         Width           =   1350
      End
      Begin VB.ComboBox ccKec 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1515
         TabIndex        =   1
         Top             =   525
         Width           =   4260
      End
      Begin MSComCtl2.DTPicker dTerbit 
         Height          =   300
         Left            =   1515
         TabIndex        =   4
         Top             =   1500
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   186843137
         CurrentDate     =   41486
      End
      Begin MSComCtl2.DTPicker dJTempo 
         Height          =   315
         Left            =   1515
         TabIndex        =   3
         Top             =   1185
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   186843137
         CurrentDate     =   41486
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total PBB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   2205
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pengurang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2910
         TabIndex        =   19
         Top             =   1860
         Width           =   780
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Terbit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   18
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Pajak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   17
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Kelurahan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   16
         Top             =   915
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Kecamatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   165
         TabIndex        =   15
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah SPPT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   1845
         Width           =   1305
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Jatuh Tempo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   13
         Top             =   1215
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmSPPT_Massal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xTT, xTB
Dim NMIN, cTarif
'Dim xMIN(2), xMAX(2)
Dim xTarif(2)
Dim cMin(2), cMax(2), cTKP(2)
Dim totChar

Private Sub cmdBangunan_Click()
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
frmOP_Tanah.Show
End Sub

Private Sub ccBayar_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub ccBayar_LostFocus()
On Error Resume Next
For i = 0 To ccBayar.ListCount - 1
        If (UCase(ccBayar.List(i)) Like "*" + UCase(ccBayar.Text) + "*" = True) Then
            ccBayar.Text = ccBayar.List(i)

            Exit Sub
        End If
          If i = ccBayar.ListCount - 1 Then
            If UCase(ccBayar.List(i)) Like "*" + UCase(ccBayar.Text) + "*" = False Then
                ccBayar.Text = ccBayar.List(0)

                Exit Sub
            End If
        End If
    Next
End Sub
Sub CALL_TBAYAR()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
ccBayar.Clear
QSTR = "SELECT * FROM TEMPAT_BAYAR ORDER BY KD_TP ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            ccBayar.AddItem rPajak!KD_TP & " " & rPajak!NM_TP
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub ccTahun_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub cmdCear_Click()
On Error Resume Next
ccTahun.Text = ccTahun.List(0)
ccKec.Text = ""
ccKel.Text = ""
dJTempo.Value = Format(Now, "dd/mm/yyyy")
dTerbit.Value = Format(Now, "dd/mm/yyyy")
tSPPT.Text = 0
tKurang.Text = 0
tTotal.Text = 0
ccBayar.Text = ccBayar.List(3)
Me.Caption = "Penetapan SPPT Massal"
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub cmdOK_Click()
On Error GoTo Salah
Dim xxPro

Screen.MousePointer = vbHourglass
Me.Caption = "Penetapan SPPT Massal"
If ccKec.Text = "" Then
    tanya1 = MsgBox("Anda Belum Memilih Wilayah Kecamatan " & _
            vbCrLf & "Seluruh Wilayah Akan diproses, Lanjut?", vbExclamation + vbYesNo, "Tetnong")
    If tanya1 = vbNo Then Screen.MousePointer = vbDefault: Exit Sub Else ccKec.Text = "*.*": ccKel.Text = "*.*": ccKel.Enabled = False
ElseIf ccKel.Text = "" Then
    tanya1 = MsgBox("Anda Belum Memilih Wilayah Kelurahan" & _
             vbCrLf & "Seluruh Kelurahan akan diproses, Lanjut?", vbExclamation + vbYesNo, "Tetnong")
    If tanya1 = vbNo Then Screen.MousePointer = vbDefault: Exit Sub Else ccKel.Text = "*.*"
End If

xSTR = "Select THN_NJOPTKP From DAT_SUBJEK_PAJAK_NJOPTKP WHERE THN_NJOPTKP='" & ccTahun.Text & "' ORDER BY THN_NJOPTKP ASC"
openDB (xSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If rPajak.EOF Then
'    MsgBox rPajak!THN_NJOPTKP
    MsgBox "NJOPTKP Untuk tahun " & ccTahun.Text & " Beluma dibuat" & _
            vbCrLf & "Proses tidak akan dilanjutkan...", vbCritical, "Tetnong"
    Screen.MousePointer = vbDefault
    Exit Sub
'rPajak.MoveNext
'Loop
End If
If ccKec.Text = "*.*" And ccKel.Text = "*.*" Then
    xSQL = "Select * From SPPT where THN_PAJAK_SPPT='" & ccTahun.Text & "'"
    Pesan1 = "SPPT seluruh/wilayah tertentu sudah ditetapkan" & _
            vbCrLf & "Anda ingin membuat ulang?"
    ccU = 1
    xxPro = "1"
    'Memindahkan SPPT lama yang sudah dimutakhirkan
ElseIf ccKec.Text <> "*.*" And ccKel.Text = "*.*" Then
    xSQL = "Select * From SPPT where  KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'"
    Pesan1 = "KEC: " & Mid(Trim(ccKec.Text), 4, Len(Trim(ccKec.Text))) & ", sudah ditetapkan" & _
            vbCrLf & "Anda ingin membuat ulang?"
            ccU = 2
            xxPro = "2"
            'Memindahkan SPPT lama yang sudah dimutakhirkan
Else
    xSQL = "Select * From SPPT where (KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(ccKel.Text), 3) & "') and THN_PAJAK_SPPT='" & ccTahun.Text & "' "
    Pesan1 = "KEC: " & Mid(ccKec.Text, 5, Len(Trim(ccKec.Text))) & "," & _
            vbCrLf & "KEL: " & Mid(ccKel.Text, 5, Len(Trim(ccKel.Text))) & ", sudah ditetapkan" & _
            vbCrLf & "Anda ingin membuat ulang?"
            ccU = 3
            xxPro = "3"
End If
'PROSES='N' BERARTI MASIH DINILAI BELUM DITETAPKAN SPPT-NYA
openDB (xSQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then
    If rPajak!PROSES = "M" Or rPajak!PROSES = "T" Then
        tanya1 = MsgBox(Pesan1, vbCritical + vbYesNo, "Tetnong: " & ccU)
        If tanya1 = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
        CPP = "M"
    Else
        tanya1 = MsgBox("Objek akan ditetapkan massal, Lanjutkan?", vbQuestion + vbYesNo, "Tetnong: " & ccU)
        If tanya1 = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
        CPP = "N"
    End If
        TANYA = MsgBox("Penetapan SPPT tunggal akan dihapus" & _
                vbCrLf & "yakin akan dilanjutkan?", vbCritical + vbYesNo, "Tetnong")
        If TANYA = vbYes Then
            If ccKec.Text = "*.*" And ccKel.Text = "*.*" Then
                xxPro = "1"
                'xSQL = "DELETE  From SPPT where THN_PAJAK_SPPT='" & ccTahun.Text & "'"
                
            ElseIf ccKec.Text <> "*.*" And ccKel.Text = "*.*" Then
                'xSQL = "DELETE  From SPPT where KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "' "
                xxPro = "2"
            Else
                xxPro = "3"
                'xSQL = "DELETE  From SPPT where KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(ccKel.Text), 3) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "' "
            End If
            'openDB (xSQL)
        Else
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
Else
        CTANYA = MsgBox("Objek pajak belum dinilai...!" & _
        vbCrLf & "Kemungkinan ada data tidak valid, Lanjutkan?", vbInformation + vbYesNo, "Tetnong..")
        If CTANYA = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

End If
pNilai.Visible = True
'call_data
'sv_SPPT
'MsgBox xxPro

        For i = 1 To 50
            pNilai.Value = i
        Next

'Menetapkan SPPT baru
 'C_STR = "iSPPT_MASSAL '" & Left(Trim(ccKec.Text), 3) & "','" & Left(Trim(ccKel.Text), 3) & "','" & ccTahun.Text & "','" & xTT & "','" & xTB & "','" & Format(dJTempo.Value, "yyyy-mm-dd") & "', '" & Round(tKurang.Text * 1, 0) & "'," & _
            "'0', '0', '0','" & Format(dTerbit.Value, "yyyy-mm-dd") & "','" & Format(dTerbit.Value, "yyyy-mm-dd") & "', '000000',1, '01', '16', '04', '01','" & Left(Trim(ccBayar.Text), 2) & "', 'M','" & xxPro & "'"
C_STR = "iSPPT_MASSAL '" & Left(Trim(ccKec.Text), 3) & "','" & Left(Trim(ccKel.Text), 3) & "','" & ccTahun.Text & "','" & xTT & "','" & xTB & "','" & Format(dJTempo.Value, "yyyy-mm-dd") & "', '" & (tKurang.Text * 1) & "'," & _
            "'0', '0', '0','" & Format(dTerbit.Value, "yyyy-mm-dd") & "','" & Format(dTerbit.Value, "yyyy-mm-dd") & "', '000000',1, '01', '16', '04', '01','" & Left(Trim(ccBayar.Text), 2) & "', 'M','" & xxPro & "'"
openDB (C_STR)
        For i = 51 To 80
            pNilai.Value = i
        Next
    
'Meng-test apakah proses pembacaan database sudah selesai...
c_sem = "Select * from SPPT"
openDB (c_sem)
        For i = 81 To 100
            pNilai.Value = i
        Next
        
strLOG = "iLOG '" & ccTahun.Text * 1 & "'"
openDB (strLOG)
MsgBox "Penetapan SPPT Massal: Sukses!"
pNilai.Visible = False
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
'
Screen.MousePointer = vbDefault
End Sub



Private Sub dJTempo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub dTerbit_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
ccTahun.Text = Format(Now, "yyyy")
ccTahun.Clear
ccTahun.Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    ccTahun.AddItem i
Next
dJTempo.Value = Format(Now, "dd/mm/yyyy")
dTerbit.Value = Format(Now, "dd/mm/yyyy")
CALL_KEC
XXSTR1 = "select  THN_AWAL_KLS_TANAH  from Kelas_tanah order by THn_AWAL_KLS_TANAH DESC"
openDB (XXSTR1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then xTT = rPajak!THN_AWAL_KLS_TANAH
XXSTR2 = "select  THN_AWAL_KLS_BNG from KELAS_BANGUNAN order by THN_AWAL_KLS_BNG DESC"
openDB (XXSTR2)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then xTB = rPajak!THN_AWAL_KLS_BNG
pNilai.Visible = False
LNilai.Visible = False
CALL_TBAYAR
ccBayar.Text = ccBayar.List(3)

End Sub
Private Sub ccTahun_LostFocus()
On Error Resume Next
For i = 0 To ccTahun.ListCount - 1
        If (UCase(ccTahun.List(i)) Like "*" + UCase(ccTahun.Text) + "*" = True) Then
            ccTahun.Text = ccTahun.List(i)
            Exit Sub
        End If
          If i = ccTahun.ListCount - 1 Then
            If UCase(ccTahun.List(i)) Like "*" + UCase(ccTahun.Text) + "*" = False Then
                ccTahun.Text = ccTahun.List(0)
                Exit Sub
            End If
        End If
    Next
End Sub
Sub CALL_KEC()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
ccKec.Clear
QSTR = "SELECT * FROM REF_KECAMATAN ORDER BY KD_KECAMATAN ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            ccKec.AddItem rPajak!KD_KECAMATAN & " " & rPajak!NM_KECAMATAN
        rPajak.MoveNext
        Loop
        ccKec.AddItem "*.*"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
Screen.MousePointer = vbDefault
End Sub
Sub CALL_KEL()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
ccKel.Clear
QSTR = "SELECT * FROM REF_KELURAHAN WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "' ORDER BY KD_KELURAHAN ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            ccKel.AddItem rPajak!KD_KELURAHAN & " " & rPajak!NM_KELURAHAN
        rPajak.MoveNext
        Loop
        ccKel.AddItem "*.*"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub

Private Sub ccKec_Click()
On Error Resume Next
If ccKec.Text = "*.*" Then
    ccKel.Enabled = False
    ccKel.Text = "*.*"

Else
    ccKel.Enabled = True
    CALL_KEL
End If
End Sub

Private Sub ccKec_KeyPress(KeyAscii As Integer)
On Error Resume Next
  If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

If InStr("0123456789*.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
         KeyAscii = 0
        End If
End Sub

Private Sub ccKec_LostFocus()

On Error Resume Next
For i = 0 To ccKec.ListCount - 1
        If (UCase(ccKec.List(i)) Like "*" + UCase(ccKec.Text) + "*" = True) Then
            ccKec.Text = ccKec.List(i)
            ccKec_Click
            Exit Sub
        End If
          If i = ccKec.ListCount - 1 Then
            If UCase(ccKec.List(i)) Like "*" + UCase(ccKec.Text) + "*" = False Then
                ccKec.Text = ccKec.List(0)
                ccKec_Click
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub ccKel_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

If InStr("0123456789*.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
         KeyAscii = 0
        End If
End Sub

Private Sub ccKel_LostFocus()
On Error Resume Next
For i = 0 To ccKel.ListCount - 1
        If (UCase(ccKel.List(i)) Like "*" + UCase(ccKel.Text) + "*" = True) Then
            ccKel.Text = ccKel.List(i)
            Exit Sub
        End If
          If i = ccKel.ListCount - 1 Then
            If UCase(ccKel.List(i)) Like "*" + UCase(ccKel.Text) + "*" = False Then
                ccKel.Text = ccKel.List(0)
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub tKurang_GotFocus()
On Error Resume Next
tKurang.SelStart = 0
tKurang.SelLength = Len(tKurang.Text)
tKurang.SetFocus
tKurang.Alignment = 0

End Sub

Private Sub tKurang_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tKurang_LostFocus()
On Error Resume Next
If tKurang.Text = "" Or tKurang.Text = "-" Or tKurang.Text = "." Then
    tKurang.Text = 0
End If
tKurang.Alignment = 1

End Sub

Private Sub tSPPT_GotFocus()
On Error Resume Next
tSPPT.SelStart = 0
tSPPT.SelLength = Len(tSPPT.Text)
tSPPT.SetFocus
tSPPT.Alignment = 0
End Sub

Private Sub tSPPT_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tSPPT_LostFocus()
On Error Resume Next
If tSPPT.Text = "" Or tSPPT.Text = "-" Or tSPPT.Text = "." Then
    tSPPT.Text = 0
End If
tSPPT.Alignment = 1

End Sub

Private Sub tTotal_GotFocus()
On Error Resume Next
tTotal.SelStart = 0
tTotal.SelLength = Len(tTotal.Text)
tTotal.SetFocus
tTotal.Alignment = 0

End Sub

Private Sub tTotal_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tTotal_LostFocus()
On Error Resume Next
If tTotal.Text = "" Or tTotal.Text = "-" Or tTotal.Text = "." Then
    tTotal.Text = 0
End If
tTotal.Alignment = 1

End Sub
Sub call_data()
On Error GoTo Salah
Dim JTotal
Screen.MousePointer = vbHourglass
pNilai.Visible = True
vOP.ListItems.Clear
    If ccKec.Text = "*.*" And ccKel.Text = "*.*" Then
        StrQ1 = "Select * From QOBJEKPAJAK ORDER BY NOPQ asc"
    ElseIf ccKec.Text <> "*.*" And ccKel.Text = "*.*" Then
        StrQ1 = "Select * From QOBJEKPAJAK WHERE KD_KECAMATAN=  '" & Left(Trim(ccKec.Text), 3) & "' ORDER BY NOPQ asc"
    Else
        StrQ1 = "Select * From QOBJEKPAJAK WHERE KD_KECAMATAN=  '" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN=  '" & Left(Trim(ccKel.Text), 3) & "'ORDER BY NOPQ asc"
    End If
    
    openDB (StrQ1)
    pNilai.Max = rPajak.RecordCount
    pNilai.Min = 1
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
    i = i + 1
        LNilai.Visible = True
        LNilai.Caption = "[1/6] Proses Pemanggilan Data Objek Pajak: " & Round(i / pNilai.Max * 100, 0) & "%"
        LNilai.Refresh
        LNilai.Visible = False
        pNilai.Value = i
        vOP.ListItems.Add i, "", Format(i, "#")
        vOP.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
        vOP.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_PROPINSI])
        vOP.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![KD_DATI2])
        vOP.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![KD_KECAMATAN])
        vOP.ListItems.Item(i).ListSubItems.Add 5, "", Trim(rPajak![KD_KELURAHAN])
        vOP.ListItems.Item(i).ListSubItems.Add 6, "", Trim(rPajak![KD_BLOK])
        vOP.ListItems.Item(i).ListSubItems.Add 7, "", Trim(rPajak![NO_URUT])
        vOP.ListItems.Item(i).ListSubItems.Add 8, "", Trim(rPajak![KD_JNS_OP])
        vOP.ListItems.Item(i).ListSubItems.Add 9, "", ccTahun.Text
        vOP.ListItems.Item(i).ListSubItems.Add 10, "", 1 'Trim(rPajak!siklus_sppt)
        vOP.ListItems.Item(i).ListSubItems.Add 11, "", "01" 'Trim(rPajak!KD_KANWIL_BANK)"
        vOP.ListItems.Item(i).ListSubItems.Add 12, "", "16" 'Trim(rPajak!KD_KPPBB_BANK)
        vOP.ListItems.Item(i).ListSubItems.Add 13, "", "04" 'Trim(rPajak!KD_BANK_TUNGGAL)
        vOP.ListItems.Item(i).ListSubItems.Add 14, "", "01" 'Trim(rPajak!KD_BANK_PERSEPSI)
        vOP.ListItems.Item(i).ListSubItems.Add 15, "", Left(Trim(ccBayar.Text), 2) '  Trim(rPajak!KD_TP)
        'vOP.ListItems.Item(i).ListSubItems.Add 16, "", Trim(rPajak!SUBJEK_PAJAK_ID)
    vOP.ListItems.Item(i).ListSubItems.Add 16, "", rPajak!Nm_wp
    vOP.ListItems.Item(i).ListSubItems.Add 17, "", rPajak!JALAN_WP
    If IsNull(rPajak!BLOK_KAV_NO_WP) = True Then rPajak!BLOK_KAV_NO_WP = "00"
    vOP.ListItems.Item(i).ListSubItems.Add 18, "", rPajak!BLOK_KAV_NO_WP
    If IsNull(rPajak!RW_WP) = True Then rPajak!RW_WP = "00"
    vOP.ListItems.Item(i).ListSubItems.Add 19, "", rPajak!RW_WP
    If IsNull(rPajak!RT_WP) = True Then rPajak!RT_WP = "00"
    vOP.ListItems.Item(i).ListSubItems.Add 20, "", rPajak!RT_WP
    If IsNull(rPajak!KELURAHAN_WP) = True Then rPajak!KELURAHAN_WP = "-"
    vOP.ListItems.Item(i).ListSubItems.Add 21, "", rPajak!KELURAHAN_WP
    If IsNull(rPajak!KOTA_WP) = True Then rPajak!KOTA_WP = "-"
    vOP.ListItems.Item(i).ListSubItems.Add 22, "", rPajak!KOTA_WP
    If IsNull(rPajak!KD_POS_WP) = True Then rPajak!KD_POS_WP = "00000"
    vOP.ListItems.Item(i).ListSubItems.Add 23, "", rPajak!KD_POS_WP
    If IsNull(rPajak!NPWP) = True Then rPajak!NPWP = "-"
    vOP.ListItems.Item(i).ListSubItems.Add 24, "", rPajak!NPWP
   '
    If IsNull(rPajak!NO_PERSIL) = True Then rPajak!NO_PERSIL = "00"
    vOP.ListItems.Item(i).ListSubItems.Add 25, "", rPajak!NO_PERSIL
    vOP.ListItems.Item(i).ListSubItems.Add 26, "", "00" 'rPajak!KD_KLS_TANAH
    vOP.ListItems.Item(i).ListSubItems.Add 27, "", xTT
    vOP.ListItems.Item(i).ListSubItems.Add 28, "", "00" 'rPajak!KD_KLS_BNG
    vOP.ListItems.Item(i).ListSubItems.Add 29, "", xTB
    vOP.ListItems.Item(i).ListSubItems.Add 30, "", Format(dJTempo.Value, "DD/MM/YYYY") ' rPajak!TGL_JATUH_TEMPO_SPPT
    vOP.ListItems.Item(i).ListSubItems.Add 31, "", Format(rPajak!TOTAL_LUAS_BUMI, "#,#0")
    vOP.ListItems.Item(i).ListSubItems.Add 32, "", Format(rPajak!TOTAL_LUAS_BNG, "#,#0")
    vOP.ListItems.Item(i).ListSubItems.Add 33, "", Format(rPajak!NJOP_BUMI, "#,#0")
    vOP.ListItems.Item(i).ListSubItems.Add 34, "", Format(rPajak!NJOP_BNG, "#,#0")
    vOP.ListItems.Item(i).ListSubItems.Add 35, "", Format(rPajak!NJOP_BUMI * 1 + rPajak!NJOP_BNG * 1, "#,#0")
    vOP.ListItems.Item(i).ListSubItems.Add 36, "", 0 'NJOPTKP
    vOP.ListItems.Item(i).ListSubItems.Add 37, "", 0 'NJKP
    vOP.ListItems.Item(i).ListSubItems.Add 38, "", 0 'HUTANG
    vOP.ListItems.Item(i).ListSubItems.Add 39, "", tKurang.Text  'PENGURANG
    vOP.ListItems.Item(i).ListSubItems.Add 40, "", 0 'JUMLAH PBB YANG BAYAR
    vOP.ListItems.Item(i).ListSubItems.Add 41, "", 0 'status pembayaran
    vOP.ListItems.Item(i).ListSubItems.Add 42, "", 0 'status penagihan
    vOP.ListItems.Item(i).ListSubItems.Add 43, "", 0 'status cetak
    vOP.ListItems.Item(i).ListSubItems.Add 44, "", Format(dTerbit.Value, "dd/mm/yyyy")
    vOP.ListItems.Item(i).ListSubItems.Add 45, "", "01/01/1900"
    vOP.ListItems.Item(i).ListSubItems.Add 46, "", "0000000000"
    vOP.ListItems.Item(i).ListSubItems.Add 47, "", "M"
    vOP.ListItems.Item(i).ListSubItems.Add 48, "", 0
    vOP.ListItems.Item(i).ListSubItems.Add 49, "", xNOP
    vOP.ListItems.Item(i).ListSubItems.Add 50, "", rPajak!JNS_BUMI
    rPajak.MoveNext
    Loop
    K_BUMI
    K_BANGUNAN
    CEK_NJOPTKP
    pNilai.Max = vOP.ListItems.Count
    pNilai.Min = 1
    For i = 1 To vOP.ListItems.Count
        LNilai.Visible = True
        LNilai.Caption = "[5/6] Hitung Nilai PBB Terhutang: " & Round(i / pNilai.Max * 100, 0) & "%"
        LNilai.Refresh
        LNilai.Visible = False
        pNilai.Value = i
        CALL_NJOPTKP
    If vOP.ListItems.Item(i).ListSubItems(35).Text * 1 >= cMin(1) And vOP.ListItems.Item(i).ListSubItems(35).Text * 1 <= cMax(1) Then
        vOP.ListItems.Item(i).ListSubItems(36).Text = Format(cTKP(1), "#,#0")
        cTarif = xTarif(1)
    Else
        vOP.ListItems.Item(i).ListSubItems(36).Text = Format(cTKP(2), "#,#0")
        cTarif = xTarif(2)
    End If
    If vOP.ListItems.Item(i).ListSubItems(48).Text = 0 Then vOP.ListItems.Item(i).ListSubItems(36).Text = 0
    'tNOP(22).Text = cTarif
    'Vop.ListItems.Item(i).ListSubItems(42).Text
    vOP.ListItems.Item(i).ListSubItems(37).Text = Format(vOP.ListItems.Item(i).ListSubItems(35).Text * 1 - vOP.ListItems.Item(i).ListSubItems(36).Text * 1, "#,#0")
    If vOP.ListItems.Item(i).ListSubItems(37).Text * 1 < 0 Then vOP.ListItems.Item(i).ListSubItems(37).Text = 0
    vOP.ListItems.Item(i).ListSubItems(38).Text = Format((vOP.ListItems.Item(i).ListSubItems(37).Text * cTarif * 1 / 100), "#,#0")
    vOP.ListItems.Item(i).ListSubItems(40).Text = Format((vOP.ListItems.Item(i).ListSubItems(38).Text * 1) - tKurang.Text * 1, "#,#0")
    
    Call_MIN
    
    If vOP.ListItems.Item(i).ListSubItems(40).Text * 1 < NMIN * 1 Then
        vOP.ListItems.Item(i).ListSubItems(40).Text = Format(NMIN, "#,#0")
    End If
    JTotal = JTotal + (vOP.ListItems.Item(i).ListSubItems(40).Text * 1)
    Next
    
    tSPPT.Text = Format(vOP.ListItems.Count, "#,#0")
    tSPPT.Refresh
    tTotal.Text = Format(JTotal, "#,#0")
    tTotal.Refresh
    'LNilai.Visible = True
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub
Sub CALL_NJOPTKP()
On Error GoTo Salah
xxSTR = "Select * From Tarif order by NJOP_MIN"
openDB (xxSTR)
'pNilai.Max = vOP.ListItems.Count
'pNilai.Min = 1
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
Do While Not rPajak.EOF
    i = i + 1
'    LNilai.Visible = True
'    LNilai.Caption = "Memanggil Tabel Tarif dan NJOPTKP"
'    LNilai.Refresh
'    LNilai.Visible = False
'    pNilai.Value = i
    cMin(i) = rPajak!NJOP_MIN
    cMax(i) = rPajak!NJOP_MAX
    cTKP(i) = rPajak!NJOPTKP
    xTarif(i) = rPajak!NILAI_TARIF
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

'Sub Call_Proses()
'On Error GoTo Salah
'    tNOP(17).Text = Format(tNOP(15).Text * 1 + tNOP(16).Text * 1, "#,#0")
'    CALL_NJOPTKP
'    If tNOP(17).Text * 1 >= cMin(1) And tNOP(17).Text * 1 <= cMax(1) Then
'        tNOP(18).Text = Format(cTKP(1), "#,#0")
'        cTarif = xTarif(1)
'    Else
'        tNOP(18).Text = Format(cTKP(1), "#,#0")
'        cTarif = xTarif(2)
'    End If
'    'If tNOP(12).Text * 1 <= 0 Or tNOP(16).Text * 1 <= 0 Then tNOP(18).Text = 0
'    CEK_NJOPTKP
'    tNOP(22).Text = cTarif
'    K_BUMI
'    K_BANGUNAN
'    tNOP(21).Text = Format(tNOP(17).Text * 1 - tNOP(18).Text * 1, "#,#0")
'    tNOP(20).Text = Format((tNOP(21).Text * tNOP(22).Text * 1 / 100) - tNOP(19).Text * 1, "#,#0")
'
'    Call_MIN
'
'    If tNOP(21).Text * 1 < 0 Then tNOP(21).Text = 0
'    If tNOP(20).Text * 1 < NMIN * 1 Then
'        tNOP(20).Text = Format(NMIN, "#,#0")
'    End If
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'Salah:
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'MsgBox Err.Number & ": " & Err.Description
'End Sub

Sub CEK_NJOPTKP()
On Error GoTo Salah
'List1.Clear
'n_STR = "select * from DAT_SUBJEK_PAJAK_NJOPTKP WHERE trim(SUBJEK_PAJAK_ID)='" & Trim(tID.Text) & "' ORDER BY SUBJEK_PAJAK_ID ASC"
    If ccKec.Text = "*.*" And ccKel.Text = "*.*" Then
        n_STR = "select * from DAT_SUBJEK_PAJAK_NJOPTKP ORDER BY SUBJEK_PAJAK_ID ASC"
    ElseIf ccKec.Text <> "*.*" And ccKel.Text = "*.*" Then
        n_STR = "select * from DAT_SUBJEK_PAJAK_NJOPTKP where KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' ORDER BY SUBJEK_PAJAK_ID ASC"
    Else
        n_STR = "select * from DAT_SUBJEK_PAJAK_NJOPTKP where KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(ccKel.Text), 3) & "' ORDER BY SUBJEK_PAJAK_ID ASC"
    End If

openDB (n_STR)
pNilai.Max = rPajak.RecordCount
pNilai.Min = 1
J = 0
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    J = J + 1
        LNilai.Visible = True
        LNilai.Caption = "[4/6] Memberi Flag NJOPTKP: " & Round(J / pNilai.Max * 100, 0) & "%"
        LNilai.Refresh
        LNilai.Visible = False
        pNilai.Value = J
    For i = 1 To vOP.ListItems.Count
        
        If (rPajak!KD_KECAMATAN = vOP.ListItems.Item(i).ListSubItems(4).Text) And (rPajak!KD_KELURAHAN = vOP.ListItems.Item(i).ListSubItems(5).Text) And (rPajak!KD_BLOK = vOP.ListItems.Item(i).ListSubItems(6).Text) And (rPajak!NO_URUT = vOP.ListItems.Item(i).ListSubItems(7).Text) And (rPajak!KD_JNS_OP = vOP.ListItems.Item(i).ListSubItems(8).Text) Then
            vOP.ListItems.Item(i).ListSubItems(48).Text = 1
        End If
    Next
    rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub K_BUMI()
On Error GoTo Salah
Dim L_BUMI
xxSTR = "select * from Kelas_TANAH WHERE THN_AWAL_KLS_TANAH='" & xTT & "'"
openDB (xxSTR)
pNilai.Max = vOP.ListItems.Count
pNilai.Min = 1
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
J = 0
Do While Not rPajak.EOF
        J = J + 1
        LNilai.Visible = True
        LNilai.Caption = "[2/6] Proses Klasifikasi Nilai Tanah: " & Round(J / pNilai.Max * 100, 0) & "%"
        LNilai.Refresh
        LNilai.Visible = False
        pNilai.Value = J

        For i = 1 To vOP.ListItems.Count
    
        L_BUMI = vOP.ListItems.Item(i).ListSubItems(31).Text * 1
        If L_BUMI <= 0 Then L_BUMI = 1
        If Format(vOP.ListItems.Item(i).ListSubItems(33).Text * 1 / L_BUMI, "#,#0") = Format(rPajak!NILAI_PER_M2_TANAH * 1000, "#,#0") Then
            vOP.ListItems.Item(i).ListSubItems(26).Text = rPajak!KD_KLS_TANAH
            'Exit Sub
            
        End If
    Next
rPajak.MoveNext
Loop
'vOP.ListItems.Item(i).ListSubItems(26).Text = "00"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub K_BANGUNAN()
On Error GoTo Salah

Dim L_BNG
xxSTR = "select * from Kelas_BANGUNAN WHERE THN_AWAL_KLS_BNG='" & xTB & "'"
openDB (xxSTR)
pNilai.Max = vOP.ListItems.Count
pNilai.Min = 1
J = 0
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
J = J + 1
        
        LNilai.Visible = True
        LNilai.Caption = "[3/6] Proses Klasifikasi Nilai Bangunan: " & Round(J / pNilai.Max * 100, 0) & "%"
        LNilai.Refresh
        LNilai.Visible = False
        pNilai.Value = J
    For i = 1 To vOP.ListItems.Count
        
        L_BNG = vOP.ListItems.Item(i).ListSubItems(32).Text * 1
        If L_BNG <= 0 Then L_BNG = 1
        If Format(vOP.ListItems.Item(i).ListSubItems(34).Text * 1 / L_BNG * 1, "#,#0") = Format(rPajak!NILAI_PER_M2_BNG * 1000, "#,#0") Then
            vOP.ListItems.Item(i).ListSubItems(28).Text = rPajak!KD_KLS_BNG
            'Exit Sub
            
        End If
    Next
rPajak.MoveNext
Loop
'vOP.ListItems.Item(i).ListSubItems(28).Text = "00"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub Call_MIN()
On Error GoTo Salah
xxSTR = "Select * From  PBB_MINIMAL WHERE THN_PBB_MINIMAL ='" & ccTahun.Text & "'order by THN_PBB_MINIMAL DESC "
openDB (xxSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'pNilai.Max = vOP.ListItems.Count
'pNilai.Min = 1
i = 0
Do While Not rPajak.EOF
i = i + 1
'    LNilai.Visible = True
'    LNilai.Caption = "Memanggil Nilai PBB Minimal"
'    LNilai.Refresh
'    pNilai.Value = i
    NMIN = rPajak!NILAI_PBB_MINIMAL
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub vOP_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vOP.SortKey = ColumnHeader.Index - 1
vOP.Sorted = True
vOP.Sorted = False
vOP.SortOrder = lvwAscending

End Sub
Sub sv_SPPT()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
Dim jumRec, JTotal
'For Each Control In Me
'If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
'    If Control.Text = "" Then
'        MsgBox "Masih ada data kosong,,,", vbCritical, "Tetnong..."
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'End If
'
'Next
'    If ccKec.Text = "" Or ccKel.Text = "" Then
'        MsgBox "Masih ada data kosong...", vbCritical, "Tetnong..."
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    If tSPPT.Text = "" Or tSPPT.Text = 0 Then
        MsgBox "Tidak Ada SPPT yang akan diproses...", vbCritical, "Tetnong..."
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    

''Cek Keberadaan NOP
'StrQ1 = "Select * From QOBJEKPAJAK WHERE NOPQ =  '" & Trim(aNOP.Text) & "' ORDER BY nopq asc"
'    openDB (StrQ1)
'    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    If rPajak.EOF Then
'        MsgBox "Data Tidak Ditemukan...", vbCritical, "Tetnong..."
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'
'xSQL = "Select * From SPPT where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & Trim(aNOP.Text) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'"
xSQL = "Select * From SPPT"
openDB (xSQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
pNilai.Max = vOP.ListItems.Count
pNilai.Min = 1
  
    
For i = 1 To vOP.ListItems.Count
'xxKec = Trim(vOP.ListItems.Item(i).ListSubItems(4).Text)
'xxKel = Trim(vOP.ListItems.Item(i).ListSubItems(5).Text)
'xxBlok = Trim(vOP.ListItems.Item(i).ListSubItems(6).Text)
'xxUrut = Trim(vOP.ListItems.Item(i).ListSubItems(7).Text)
'xxJenis = Trim(vOP.ListItems.Item(i).ListSubItems(8).Text)
'xHutang = Format(vOP.ListItems.Item(i).ListSubItems(38).Text, "#,#0")
'xKurang = Format(tKurang.Text * 1, "#,#0")
'xBayar = Format(vOP.ListItems.Item(i).ListSubItems(40).Text * 1, "#,#0")
'xxTerbit = Format(dTerbit.Value, "dd/mm/yyyy")
'xxJTempo = Format(dJTempo.Value, "dd/mm/yyyy")
'xxNama = vOP.ListItems.Item(i).ListSubItems(16).Text
'xxAlamat = vOP.ListItems.Item(i).ListSubItems(17).Text
'xxBlok = vOP.ListItems.Item(i).ListSubItems(18).Text
'xxRW = vOP.ListItems.Item(i).ListSubItems(19).Text
'xxRT = vOP.ListItems.Item(i).ListSubItems(20).Text
'xxLurah = vOP.ListItems.Item(i).ListSubItems(21).Text
'xxKota = vOP.ListItems.Item(i).ListSubItems(22).Text
'xxPos = vOP.ListItems.Item(i).ListSubItems(23).Text
'xxNPWP = vOP.ListItems.Item(i).ListSubItems(24).Text
'xxPersil = vOP.ListItems.Item(i).ListSubItems(25).Text
If vOP.ListItems.Item(i).ListSubItems(50).Text <> "4" Then
        LNilai.Visible = True
        LNilai.Caption = "[6/6] Proses Pembentukan Database SPPT: " & Round(i / pNilai.Max * 100, 0) & "%"
        LNilai.Refresh
        LNilai.Visible = False
        pNilai.Visible = True
        pNilai.Value = i
    
'iSQL = "INSERT INTO SPPT(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,THN_PAJAK_SPPT,NM_WP_SPPT,JLN_WP_SPPT,BLOK_KAV_NO_WP_SPPT,RW_WP_SPPT,RT_WP_SPPT,KELURAHAN_WP_SPPT,KOTA_WP_SPPT,KD_POS_WP_SPPT,NPWP_SPPT," & _
    "NO_PERSIL_SPPT,KD_KLS_TANAH,THN_AWAL_KLS_TANAH,KD_KLS_BNG,THN_AWAL_KLS_BNG,TGL_JATUH_TEMPO_SPPT,LUAS_BUMI_SPPT,LUAS_BNG_SPPT, NJOP_BUMI_SPPT,NJOP_BNG_SPPT,NJOP_SPPT,NJOPTKP_SPPT,NJKP_SPPT,PBB_TERHUTANG_SPPT,FAKTOR_PENGURANG_SPPT," & _
    "PBB_YG_HARUS_DIBAYAR_SPPT,STATUS_PEMBAYARAN_SPPT,STATUS_TAGIHAN_SPPT,STATUS_CETAK_SPPT,TGL_TERBIT_SPPT,TGL_CETAK_SPPT,NIP_PENCETAK_SPPT,SIKLUS_SPPT,KD_KANWIL_BANK,KD_KPPBB_BANK,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,KD_TP,PROSES)" & _
    "Values('" & Trim(vOP.ListItems.Item(i).ListSubItems(2).Text) & "', '" & Trim(vOP.ListItems.Item(i).ListSubItems(3).Text) & "', '" & Trim(vOP.ListItems.Item(i).ListSubItems(4).Text) & "', '" & Trim(vOP.ListItems.Item(i).ListSubItems(5).Text) & "', '" & Trim(vOP.ListItems.Item(i).ListSubItems(6).Text) & "', '" & Trim(vOP.ListItems.Item(i).ListSubItems(7).Text) & "', '" & Trim(vOP.ListItems.Item(i).ListSubItems(8).Text) & "', '" & ccTahun.Text & "'," & _
    "'" & Trim(vOP.ListItems.Item(i).ListSubItems(16).Text) & "', '" & Trim(vOP.ListItems.Item(i).ListSubItems(17).Text) & "', '" & Trim(vOP.ListItems.Item(i).ListSubItems(18).Text) & "', '" & Trim(vOP.ListItems.Item(i).ListSubItems(19).Text) & "', '" & Trim(vOP.ListItems.Item(i).ListSubItems(20).Text) & "','" & Trim(vOP.ListItems.Item(i).ListSubItems(21).Text) & "','" & Trim(vOP.ListItems.Item(i).ListSubItems(22).Text) & "','" & Trim(vOP.ListItems.Item(i).ListSubItems(23).Text) & "','" & Trim(vOP.ListItems.Item(i).ListSubItems(24).Text) & "','" & Trim(vOP.ListItems.Item(i).ListSubItems(25).Text) & "'," & _
    " '" & vOP.ListItems.Item(i).ListSubItems(26).Text & " ','" & vOP.ListItems.Item(i).ListSubItems(27).Text & "','" & vOP.ListItems.Item(i).ListSubItems(28).Text & "','" & vOP.ListItems.Item(i).ListSubItems(29).Text & "','" & vOP.ListItems.Item(i).ListSubItems(30).Text & "','" & vOP.ListItems.Item(i).ListSubItems(31).Text & "','" & vOP.ListItems.Item(i).ListSubItems(32).Text & "','" & vOP.ListItems.Item(i).ListSubItems(33).Text & "','" & vOP.ListItems.Item(i).ListSubItems(34).Text & "','" & vOP.ListItems.Item(i).ListSubItems(35).Text & "'," & _
    "'" & vOP.ListItems.Item(i).ListSubItems(36).Text & "','" & vOP.ListItems.Item(i).ListSubItems(37).Text & "','" & Trim(vOP.ListItems.Item(i).ListSubItems(38).Text) & "','" & Trim(vOP.ListItems.Item(i).ListSubItems(39).Text) & "','" & Trim(vOP.ListItems.Item(i).ListSubItems(40).Text) & "','" & vOP.ListItems.Item(i).ListSubItems(41).Text & "','" & vOP.ListItems.Item(i).ListSubItems(42).Text & "','" & vOP.ListItems.Item(i).ListSubItems(43).Text & "','" & vOP.ListItems.Item(i).ListSubItems(44).Text & "','" & vOP.ListItems.Item(i).ListSubItems(45).Text & "'," & _
    "'" & vOP.ListItems.Item(i).ListSubItems(46).Text & "', '" & vOP.ListItems.Item(i).ListSubItems(10).Text & "' ,'" & vOP.ListItems.Item(i).ListSubItems(11).Text & "','" & vOP.ListItems.Item(i).ListSubItems(12).Text & "','" & vOP.ListItems.Item(i).ListSubItems(13).Text & "','" & vOP.ListItems.Item(i).ListSubItems(14).Text & "','" & vOP.ListItems.Item(i).ListSubItems(15).Text & "','M')"
 '   openDB (iSQL)
    
'rPajak.AddNew
'rPajak!KD_PROPINSI = "12"
'rPajak!KD_DATI2 = "12"
'rPajak!KD_KECAMATAN = xxKec
'rPajak!KD_KELURAHAN = xxKel
'rPajak!KD_BLOK = xxBlok
'rPajak!NO_URUT = xxUrut
'rPajak!KD_JNS_OP = xxJenis
'rPajak!THN_PAJAK_SPPT = ccTahun.Text
'rPajak!NM_WP_SPPT = xxNama
'rPajak!JLN_WP_SPPT = xxAlamat
'rPajak!BLOK_KAV_NO_WP_SPPT = xxBlok
'rPajak!RW_WP_SPPT = xxRW
'rPajak!RT_WP_SPPT = xxRT
'rPajak!KELURAHAN_WP_SPPT = xxLurah
'rPajak!KOTA_WP_SPPT = xxKota
'rPajak!KD_POS_WP_SPPT = xxPos
'rPajak!NPWP_SPPT = xxNPWP
'rPajak!NO_PERSIL_SPPT = xxPersil
'rPajak!KD_KLS_TANAH = vOP.ListItems.Item(i).ListSubItems(26).Text
'rPajak!THN_AWAL_KLS_TANAH = xTT 'vOP.ListItems.Item(i).ListSubItems(27).Text
'rPajak!KD_KLS_BNG = vOP.ListItems.Item(i).ListSubItems(28).Text
'rPajak!THN_AWAL_KLS_BNG = xTB 'vOP.ListItems.Item(i).ListSubItems(29).Text
'rPajak!TGL_JATUH_TEMPO_SPPT = Format(dJTempo.Value, "DD/MM/YYYY") 'vOP.ListItems.Item(i).ListSubItems(30).Text
'rPajak!LUAS_BUMI_SPPT = vOP.ListItems.Item(i).ListSubItems(31).Text
'rPajak!LUAS_BNG_SPPT = vOP.ListItems.Item(i).ListSubItems(32).Text
'rPajak!NJOP_BUMI_SPPT = vOP.ListItems.Item(i).ListSubItems(33).Text
'rPajak!NJOP_BNG_SPPT = vOP.ListItems.Item(i).ListSubItems(34).Text
'rPajak!NJOP_SPPT = vOP.ListItems.Item(i).ListSubItems(35).Text
'rPajak!NJOPTKP_SPPT = vOP.ListItems.Item(i).ListSubItems(36).Text
'rPajak!NJKP_SPPT = vOP.ListItems.Item(i).ListSubItems(37).Text
'rPajak!PBB_TERHUTANG_SPPT = xHutang
'rPajak!FAKTOR_PENGURANG_SPPT = xKurang
'rPajak!PBB_YG_HARUS_DIBAYAR_SPPT = xBayar
'rPajak!STATUS_PEMBAYARAN_SPPT = vOP.ListItems.Item(i).ListSubItems(41).Text
'rPajak!STATUS_TAGIHAN_SPPT = vOP.ListItems.Item(i).ListSubItems(42).Text
'rPajak!STATUS_CETAK_SPPT = vOP.ListItems.Item(i).ListSubItems(43).Text
'rPajak!TGL_TERBIT_SPPT = Format(dTerbit.Value, "DD/MM/YYYY") 'vOP.ListItems.Item(i).ListSubItems(44).Text
'rPajak!TGL_CETAK_SPPT = vOP.ListItems.Item(i).ListSubItems(45).Text
'rPajak!NIP_PENCETAK_SPPT = vOP.ListItems.Item(i).ListSubItems(46).Text
'rPajak!SIKLUS_SPPT = vOP.ListItems.Item(i).ListSubItems(10).Text
'rPajak!KD_KANWIL_BANK = vOP.ListItems.Item(i).ListSubItems(11).Text
'rPajak!KD_KPPBB_BANK = vOP.ListItems.Item(i).ListSubItems(12).Text
'rPajak!KD_BANK_TUNGGAL = vOP.ListItems.Item(i).ListSubItems(13).Text
'rPajak!KD_BANK_PERSEPSI = vOP.ListItems.Item(i).ListSubItems(14).Text
'rPajak!KD_TP = vOP.ListItems.Item(i).ListSubItems(15).Text
'rPajak!PROSES = "M"
'rPajak.Update
rPajak.AddNew
rPajak!KD_PROPINSI = Trim(vOP.ListItems.Item(i).ListSubItems(2).Text)
rPajak!KD_DATI2 = Trim(vOP.ListItems.Item(i).ListSubItems(3).Text)
rPajak!KD_KECAMATAN = Trim(vOP.ListItems.Item(i).ListSubItems(4).Text)
rPajak!KD_KELURAHAN = Trim(vOP.ListItems.Item(i).ListSubItems(5).Text)
rPajak!KD_BLOK = Trim(vOP.ListItems.Item(i).ListSubItems(6).Text)
rPajak!NO_URUT = Trim(vOP.ListItems.Item(i).ListSubItems(7).Text)
rPajak!KD_JNS_OP = Trim(vOP.ListItems.Item(i).ListSubItems(8).Text)
rPajak!THN_PAJAK_SPPT = ccTahun.Text
rPajak!NM_WP_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(16).Text)
rPajak!JLN_WP_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(17).Text)
rPajak!BLOK_KAV_NO_WP_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(18).Text)
rPajak!RW_WP_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(19).Text)
rPajak!RT_WP_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(20).Text)
rPajak!KELURAHAN_WP_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(21).Text)
rPajak!KOTA_WP_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(22).Text)
rPajak!KD_POS_WP_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(23).Text)
rPajak!NPWP_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(24).Text)
rPajak!NO_PERSIL_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(25).Text)
rPajak!KD_KLS_TANAH = vOP.ListItems.Item(i).ListSubItems(26).Text
rPajak!THN_AWAL_KLS_TANAH = vOP.ListItems.Item(i).ListSubItems(27).Text
rPajak!KD_KLS_BNG = vOP.ListItems.Item(i).ListSubItems(28).Text
rPajak!THN_AWAL_KLS_BNG = vOP.ListItems.Item(i).ListSubItems(29).Text
rPajak!TGL_JATUH_TEMPO_SPPT = vOP.ListItems.Item(i).ListSubItems(30).Text
rPajak!LUAS_BUMI_SPPT = vOP.ListItems.Item(i).ListSubItems(31).Text
rPajak!LUAS_BNG_SPPT = vOP.ListItems.Item(i).ListSubItems(32).Text
rPajak!NJOP_BUMI_SPPT = vOP.ListItems.Item(i).ListSubItems(33).Text
rPajak!NJOP_BNG_SPPT = vOP.ListItems.Item(i).ListSubItems(34).Text
rPajak!NJOP_SPPT = vOP.ListItems.Item(i).ListSubItems(35).Text
rPajak!NJOPTKP_SPPT = vOP.ListItems.Item(i).ListSubItems(36).Text
rPajak!NJKP_SPPT = vOP.ListItems.Item(i).ListSubItems(37).Text
rPajak!PBB_TERHUTANG_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(38).Text)
rPajak!FAKTOR_PENGURANG_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(39).Text)
rPajak!PBB_YG_HARUS_DIBAYAR_SPPT = Trim(vOP.ListItems.Item(i).ListSubItems(40).Text)
rPajak!STATUS_PEMBAYARAN_SPPT = vOP.ListItems.Item(i).ListSubItems(41).Text
rPajak!STATUS_TAGIHAN_SPPT = vOP.ListItems.Item(i).ListSubItems(42).Text
rPajak!STATUS_CETAK_SPPT = vOP.ListItems.Item(i).ListSubItems(43).Text
rPajak!TGL_TERBIT_SPPT = vOP.ListItems.Item(i).ListSubItems(44).Text
rPajak!TGL_CETAK_SPPT = vOP.ListItems.Item(i).ListSubItems(45).Text
rPajak!NIP_PENCETAK_SPPT = vOP.ListItems.Item(i).ListSubItems(46).Text
rPajak!SIKLUS_SPPT = vOP.ListItems.Item(i).ListSubItems(10).Text
rPajak!KD_KANWIL_BANK = vOP.ListItems.Item(i).ListSubItems(11).Text
rPajak!KD_KPPBB_BANK = vOP.ListItems.Item(i).ListSubItems(12).Text
rPajak!KD_BANK_TUNGGAL = vOP.ListItems.Item(i).ListSubItems(13).Text
rPajak!KD_BANK_PERSEPSI = vOP.ListItems.Item(i).ListSubItems(14).Text
rPajak!KD_TP = Left(Trim(ccBayar.Text), 2) 'vOP.ListItems.Item(i).ListSubItems(15).Text
rPajak!PROSES = "M"
rPajak.Update
JTotal = JTotal + (vOP.ListItems.Item(i).ListSubItems(40).Text * 1)
    jumRec = jumRec + 1
End If

Next
    tSPPT.Text = Format(jumRec, "#,#0")
    tSPPT.Refresh
    tTotal.Text = Format(JTotal, "#,#0")
    tTotal.Refresh
    Me.Caption = "Penetapan SPPT Massal: Sukses!"
    pNilai.Visible = False
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub


