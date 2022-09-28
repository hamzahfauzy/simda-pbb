VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBayar1 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pembayaran/Pelunasan SPPT Secara Tunggal"
   ClientHeight    =   4710
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   5820
   ControlBox      =   0   'False
   Icon            =   "frmBayar1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5820
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   6150
      TabIndex        =   24
      Top             =   0
      Width           =   6150
      Begin VB.CheckBox hTunggal 
         BackColor       =   &H80000002&
         Caption         =   "Hapus Pelunasan"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   33
         Top             =   105
         Width           =   3270
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   -15
      TabIndex        =   29
      Top             =   405
      Width           =   5925
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
         Left            =   1500
         TabIndex        =   0
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun  Pajak"
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
         Left            =   135
         TabIndex        =   30
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      Height          =   510
      Left            =   -30
      TabIndex        =   27
      Top             =   765
      Width           =   5895
      Begin VB.CommandButton cmdNOP1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5475
         TabIndex        =   32
         Top             =   150
         Width           =   345
      End
      Begin MSMask.MaskEdBox aNOP 
         Height          =   315
         Left            =   1485
         TabIndex        =   1
         Top             =   150
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox tNOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   1515
         TabIndex        =   31
         Top             =   150
         Width           =   3870
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "NOP"
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
         Left            =   195
         TabIndex        =   28
         Top             =   195
         Width           =   1215
      End
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
      Height          =   2700
      Left            =   -15
      TabIndex        =   14
      Top             =   825
      Width           =   5895
      Begin VB.TextBox tNo 
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
         TabIndex        =   2
         Text            =   "1"
         Top             =   525
         Width           =   1275
      End
      Begin VB.TextBox tDenda 
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
         TabIndex        =   4
         Text            =   "0"
         Top             =   1200
         Width           =   2565
      End
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
         Left            =   1515
         TabIndex        =   7
         Top             =   2220
         Width           =   4260
      End
      Begin VB.TextBox tPBB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   5
         Text            =   "0"
         Top             =   1545
         Width           =   2565
      End
      Begin VB.TextBox tSPPT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
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
         TabIndex        =   3
         Text            =   "0"
         Top             =   855
         Width           =   2565
      End
      Begin MSComCtl2.DTPicker dBayar 
         Height          =   315
         Left            =   1515
         TabIndex        =   6
         Top             =   1890
         Width           =   1620
         _ExtentX        =   2858
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
         Format          =   152240129
         CurrentDate     =   41486
      End
      Begin MSComCtl2.UpDown xUP 
         Height          =   315
         Left            =   2805
         TabIndex        =   11
         Top             =   510
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Denda"
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
         TabIndex        =   23
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Left            =   180
         TabIndex        =   22
         Top             =   2265
         Width           =   1305
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Jlh Dibayarkan"
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
         Left            =   180
         TabIndex        =   21
         Top             =   1605
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah PBB"
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
         TabIndex        =   17
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "SPPT Ke-"
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
         Left            =   180
         TabIndex        =   16
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Left            =   180
         TabIndex        =   15
         Top             =   1935
         Width           =   1305
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   -15
      TabIndex        =   18
      Top             =   3420
      Width           =   5895
      Begin VB.TextBox tNIP 
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
         Left            =   3210
         TabIndex        =   9
         Text            =   "0"
         Top             =   210
         Width           =   2565
      End
      Begin MSComCtl2.DTPicker dRekam 
         Height          =   315
         Left            =   1515
         TabIndex        =   8
         Top             =   225
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   152240129
         CurrentDate     =   41486
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Perekaman"
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
         TabIndex        =   20
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIP"
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
         Left            =   2865
         TabIndex        =   19
         Top             =   255
         Width           =   255
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
      TabIndex        =   13
      Top             =   4230
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
      Left            =   2340
      TabIndex        =   12
      Top             =   4230
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
      TabIndex        =   10
      Top             =   4230
      Width           =   915
   End
   Begin MSComctlLib.ListView vOP 
      Height          =   4245
      Left            =   6120
      TabIndex        =   26
      Top             =   15
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   7488
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
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   -105
      ScaleHeight     =   720
      ScaleWidth      =   6150
      TabIndex        =   25
      Top             =   4050
      Width           =   6150
   End
End
Attribute VB_Name = "frmBayar1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totChar, JSISA, S_BAYAR
Private Sub cmdBangunan_Click()
On Error Resume Next
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
On Error Resume Next
frmOP_Tanah.Show
End Sub

Private Sub aNOP_Change()
On Error Resume Next
tNOP(0).Text = aNOP.Text
End Sub

Private Sub aNOP_GotFocus()
On Error Resume Next
aNOP.Mask = "12.12.###.###.###-####.#"

End Sub

Private Sub aNOP_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub aNOP_LostFocus()
On Error Resume Next
call_data
    tNOP(0).Text = aNOP.Text

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
On Error GoTo Salah
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
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

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


Private Sub cmdClear_Click()
On Error Resume Next
ccTahun.Text = ccTahun.List(0)
dBayar.Value = Format(Now, "dd/mm/yyyy")
dRekam.Value = Format(Now, "dd/mm/yyyy")
tSPPT.Text = 0
tPBB.Text = 0
tNIP.Text = 0
tDenda.Text = 0
tNo.Text = 1
ccBayar.Text = ccBayar.List(3)
hTunggal.Value = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNOP1_Click()
On Error Resume Next
J_Karakter
If Len(Trim(tNOP(0).Text)) - (totChar * 1) = 24 Then
    call_data
Else
    xID = 5
    frmLIST_Objek1.Show
End If
If hTunggal.Value = 1 Then cUbah
End Sub

Private Sub cmdOK_Click()
On Error GoTo Salah
    'Cek Apakah sudah dilakukan penetapan SPPT atau belum
'    s_YAR = "select * From SPPT where KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= '" & aNOP.Text & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "' " 'and (PROSES='M' or PROSES='T')"
'    openDB (s_YAR)
'    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    If Not rPajak.EOF Then
'         MsgBox "SPPT Belum ditetapkan" & _
'         vbCrLf & "proses tidak dapat dilanjutkan", vbCritical, "Tetnong..!"
'         Exit Sub
'    End If
'    'Cek Pembayaran
'call_data
'    t_YAR = "select * From SPPT where KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= '" & aNOP.Text & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "' "
'    openDB (t_YAR)
'    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    If Not rPajak.EOF Then
'        JSISA = rPajak!PBB_YG_HARUS_DIBAYAR_SPPT
'    End If
    d_YAR = "select * From PEMBAYARAN_SPPT where KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= '" & aNOP.Text & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'" ' and PEMBAYARAN_SPPT_KE='" & tNo.Text & "'"
    openDB (d_YAR)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If hTunggal.Value = 0 Then
    Do While Not rPajak.EOF
        If rPajak!PEMBAYARAN_SPPT_KE = tNo.Text * 1 Then MsgBox "Silahkan masukkan termin SPPT Ke-", vbCritical, "Error": GoTo Salah
        rPajak.MoveNext
    Loop
    End If
    If Not rPajak.EOF Then
        If hTunggal.Value = 0 Then
        
            If JSISA * 1 <= 0 Then
            TANYA = MsgBox("NOP : " & aNOP.Text & ", sudah lunas" & _
                  vbCrLf & "Apa anda ingin mengulangi?", vbInformation + vbYesNo, "Exist...!")
                  
            If TANYA = vbNo Then
               ' Hapus_BYR
               Exit Sub
            Else
                C_STR = "HAPUS_LUNAS_TUNGGAL '" & ccTahun.Text & "','" & tNo.Text * 1 & "','" & aNOP.Text & "'"
                openDB (C_STR)
                
            End If
            End If
        Else
            
            CTANYA = MsgBox("Yakin Hapus Objek Sudah Lunas?", vbQuestion + vbYesNo, "Deleted...!")
            If CTANYA = vbNo Then Exit Sub
                C_STR = "HAPUS_LUNAS_TUNGGAL '" & ccTahun.Text & "','" & tNo.Text * 1 & "','" & aNOP.Text & "'"
                openDB (C_STR)
                GoTo Keluar
        End If
    
    End If
'sv_Bayar

If tPBB.Text * 1 > JSISA And tDenda.Text <= 0 And tSPPT.Text * 1 = 0 Then MsgBox "Melebihi Pajak Terhutang...", vbCritical + vbOKOnly, "Error": Exit Sub
If tPBB.Text * 1 <= 0 Then MsgBox "Tidak boleh Nol...", vbCritical + vbOKOnly, "Error": Exit Sub
    C_STR = "LUNAS_TUNGGAL '" & ccTahun.Text & "','" & tNo.Text * 1 & "','" & Left(Trim(ccBayar.Text), 2) & "','" & Round(tDenda.Text, 0) & "','" & Format(dBayar.Value, "yyyy-mm-dd") & "','" & Format(dRekam.Value, "yyyy-mm-dd") & "', '" & tNIP.Text & "','" & aNOP.Text & "','" & tPBB.Text * 1 & "'"
    openDB (C_STR)
Keluar:
cmdClear_Click
MsgBox "Proses berhasil...!", vbExclamation, "Sukses!"
hTunggal.Value = 0
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description


End Sub

Private Sub dBayar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
End If
End Sub

Private Sub dRekam_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
End If

End Sub

Private Sub Form_Activate()
On Error GoTo Salah
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
    ccTahun.Text = Format(Now, "yyyy")
ccTahun.Clear
ccTahun.Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    ccTahun.AddItem i
Next
dBayar.Value = Format(Now, "dd/mm/yyyy")
dRekam.Value = Format(Now, "dd/mm/yyyy")
CALL_TBAYAR
ccBayar.Text = ccBayar.List(3)
hTunggal.Value = 0
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description
Keluar:

End Sub

Private Sub ccTahun_LostFocus()
On Error GoTo Salah
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
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

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
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:
Screen.MousePointer = vbDefault
End Sub

Private Sub hTunggal_Click()
On Error Resume Next
If hTunggal.Value = 1 Then
    cUbah
End If
End Sub

Private Sub tDenda_GotFocus()
On Error Resume Next
tDenda.SelStart = 0
tDenda.SelLength = Len(tDenda.Text)
tDenda.SetFocus
tDenda.Alignment = 0

End Sub

Private Sub tDenda_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tDenda_LostFocus()
On Error Resume Next

tDenda.Alignment = 1
If tPBB.Text = "" Or tPBB.Text = "." Or tPBB.Text = "," Or tPBB.Text = "-" Then tPBB.Text = 0
If tDenda.Text = "" Or tDenda.Text = "." Or tDenda.Text = "," Or tDenda.Text = "-" Then tDenda.Text = 0
tPBB.Text = Format(tDenda.Text * 1 + tSPPT.Text * 1, "#,#0")
vOP.ListItems.Item(1).ListSubItems(16).Text = Format(tDenda.Text, "#,#0")
vOP.ListItems.Item(1).ListSubItems(17).Text = Format(tPBB.Text, "#,#0")
tDenda.Text = Format(tDenda.Text, "#,#0")
'cmdNOP1_Click
End Sub

Private Sub tNIP_GotFocus()
On Error Resume Next
tNIP.SelStart = 0
tNIP.SelLength = Len(tNIP.Text)
tNIP.SetFocus
tNIP.Alignment = 0
End Sub

Private Sub tNIP_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tNIP_LostFocus()
On Error Resume Next
If tNIP.Text = "" Or tNIP.Text = "." Or tNIP.Text = "," Or tNIP.Text = "-" Then tNIP.Text = 0
tNIP.Alignment = 1
End Sub
Sub sv_Bayar()
On Error GoTo Salah
If tSPPT.Text * 1 = 0 And tDenda.Text * 1 = 0 And tPBB.Text * 1 = 0 Then MsgBox "Kemungkinan ada kesalahan perhitungan", vbCritical, "Error": Exit Sub
For i = 1 To vOP.ListItems.Count
    I_yar = "INSERT INTO PEMBAYARAN_SPPT(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,THN_PAJAK_SPPT,PEMBAYARAN_SPPT_KE,KD_KANWIL_BANK, KD_KPPBB_BANK,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,KD_TP,DENDA_SPPT,JML_SPPT_YG_DIBAYAR,TGL_PEMBAYARAN_SPPT,TGL_REKAM_BYR_SPPT,NIP_REKAM_BYR_SPPT)" & _
            "VALUES ('" & vOP.ListItems.Item(i).ListSubItems(2).Text & "','" & vOP.ListItems.Item(i).ListSubItems(3).Text & "','" & vOP.ListItems.Item(i).ListSubItems(4).Text & "','" & vOP.ListItems.Item(i).ListSubItems(5).Text & "','" & vOP.ListItems.Item(i).ListSubItems(6).Text & "','" & vOP.ListItems.Item(i).ListSubItems(7).Text & "','" & vOP.ListItems.Item(i).ListSubItems(8).Text & "'," & _
            "'" & vOP.ListItems.Item(i).ListSubItems(9).Text & "','" & vOP.ListItems.Item(i).ListSubItems(10).Text & "','" & vOP.ListItems.Item(i).ListSubItems(11).Text & "','" & vOP.ListItems.Item(i).ListSubItems(12).Text & "','" & vOP.ListItems.Item(i).ListSubItems(13).Text & "','" & vOP.ListItems.Item(i).ListSubItems(14).Text & "','" & vOP.ListItems.Item(i).ListSubItems(15).Text & "'," & _
            "'" & Round(vOP.ListItems.Item(i).ListSubItems(16).Text, 0) & "','" & Round(vOP.ListItems.Item(i).ListSubItems(17).Text, 0) & "','" & Format(vOP.ListItems.Item(i).ListSubItems(18).Text, "yyyy-mm-dd") & "','" & Format(vOP.ListItems.Item(i).ListSubItems(19).Text, "yyyy-mm-dd") & "','" & vOP.ListItems.Item(i).ListSubItems(20).Text & "')"
    openDB (I_yar)
Next
MsgBox "Proses berhasil...!", vbExclamation, "Sukses!"
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub call_data()
On Error GoTo Salah
Dim JTotal, jSPPT
Screen.MousePointer = vbHourglass
vOP.ListItems.Clear
    'D_YAR = "select * From PEMBAYARAN_SPPT where KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= '" & aNOP.Text & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "' "
    d_YAR = "SELECT KD_PROPINSI, KD_DATI2, KD_KECAMATAN, KD_KELURAHAN, KD_BLOK, NO_URUT, KD_JNS_OP, THN_PAJAK_SPPT, KD_KANWIL_BANK, KD_KPPBB_BANK, KD_BANK_TUNGGAL, KD_BANK_PERSEPSI, KD_TP, sum(DENDA_SPPT) as TOT_DENDA, Sum(JML_SPPT_YG_DIBAYAR) AS TOT_BAYAR, TGL_PEMBAYARAN_SPPT, TGL_REKAM_BYR_SPPT, NIP_REKAM_BYR_SPPT" & _
            " From PEMBAYARAN_SPPT GROUP BY KD_PROPINSI, KD_DATI2, KD_KECAMATAN, KD_KELURAHAN, KD_BLOK, NO_URUT, KD_JNS_OP, THN_PAJAK_SPPT, KD_KANWIL_BANK, KD_KPPBB_BANK, KD_BANK_TUNGGAL, KD_BANK_PERSEPSI, KD_TP, TGL_PEMBAYARAN_SPPT, TGL_REKAM_BYR_SPPT, NIP_REKAM_BYR_SPPT, [KD_PROPINSI]+'.'+[KD_DATI2]+'.'+[KD_KECAMATAN]+'.'+[KD_KELURAHAN]+'.'+[KD_BLOK]+'-'+[NO_URUT]+'.'+[KD_JNS_OP] " & _
            " HAVING (((THN_PAJAK_SPPT)='" & ccTahun.Text & "') AND (([KD_PROPINSI]+'.'+[KD_DATI2]+'.'+[KD_KECAMATAN]+'.'+[KD_KELURAHAN]+'.'+[KD_BLOK]+'-'+[NO_URUT]+'.'+[KD_JNS_OP])='" & aNOP.Text & "' ))"
    openDB (d_YAR)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If Not rPajak.EOF Then
        S_BAYAR = rPajak!TOT_BAYAR
    Else
        S_BAYAR = 0
    End If
    
    StrQ1 = "Select * From SPPT WHERE KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= '" & aNOP.Text & "' and (PROSES='M' OR PROSES='T')and THN_PAJAK_SPPT='" & ccTahun.Text & "'"
    openDB (StrQ1)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        MsgBox "SPPT Belum DITETAPKAN...!", vbCritical, "Tetnong"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    i = 0
    Do While Not rPajak.EOF
        JSISA = rPajak!PBB_YG_HARUS_DIBAYAR_SPPT * 1 - S_BAYAR * 1
    i = i + 1
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
        vOP.ListItems.Item(i).ListSubItems.Add 10, "", tNo.Text
        vOP.ListItems.Item(i).ListSubItems.Add 11, "", "01" 'Trim(rPajak!KD_KANWIL_BANK)"
        vOP.ListItems.Item(i).ListSubItems.Add 12, "", "16" 'Trim(rPajak!KD_KPPBB_BANK)
        vOP.ListItems.Item(i).ListSubItems.Add 13, "", "01" 'Trim(rPajak!KD_BANK_TUNGGAL)
        vOP.ListItems.Item(i).ListSubItems.Add 14, "", "01" 'Trim(rPajak!KD_BANK_PERSEPSI)
        vOP.ListItems.Item(i).ListSubItems.Add 15, "", Left(Trim(ccBayar.Text), 2)
        vOP.ListItems.Item(i).ListSubItems.Add 16, "", Format(tDenda.Text, "#,#0.00")
        
        tSPPT.Text = Format(rPajak!PBB_YG_HARUS_DIBAYAR_SPPT * 1, "#,#0.00")
        'tPBB.Text = Format(tDenda.Text * 1 + rPajak!PBB_YG_HARUS_DIBAYAR_SPPT * 1, "#,#0.00")
        If rPajak!PBB_YG_HARUS_DIBAYAR_SPPT * 1 - S_BAYAR <= 0 Then tSPPT.Text = 0
        tPBB.Text = Format(tDenda.Text * 1 + tSPPT.Text * 1 - S_BAYAR * 1, "#,#0.00")
        If tPBB.Text <= 0 Then tPBB.Text = 0
        vOP.ListItems.Item(i).ListSubItems.Add 17, "", Format(tPBB.Text, "#,#0.00")
        vOP.ListItems.Item(i).ListSubItems.Add 18, "", Format(dBayar.Value, "DD/MM/YYYY")
        vOP.ListItems.Item(i).ListSubItems.Add 19, "", Format(dRekam.Value, "DD/MM/YYYY")
        vOP.ListItems.Item(i).ListSubItems.Add 20, "", tNIP.Text
    rPajak.MoveNext
    Loop
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:
    
Screen.MousePointer = vbDefault
End Sub

Sub Hapus_BYR()
On Error GoTo Salah
        d_YAR = "DELETE  From PEMBAYARAN_SPPT where KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= '" & aNOP.Text & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "' AND PEMBAYARAN_SPPT_KE='" & tNo.Text & "'"
        openDB (d_YAR)
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub J_Karakter()
On Error GoTo Salah
Dim jmlText, jmlChar, i As Integer
    jmlChar = 0
    jmlText = Len(tNOP(0).Text)
    For i = 0 To jmlText
        tNOP(0).SelStart = i
        tNOP(0).SelLength = 1
        If tNOP(0).SelText = "_" Then
            jmlChar = jmlChar + 1
        End If
    Next
    totChar = jmlChar
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub tNo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tPBB_GotFocus()
On Error Resume Next
tPBB.SelStart = 0
tPBB.SelLength = Len(tPBB.Text)
tPBB.SetFocus
tPBB.Alignment = 0
End Sub

Private Sub tPBB_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tPBB_LostFocus()
On Error Resume Next
tPBB.Alignment = 1
If tPBB.Text = "" Or tPBB.Text = "." Or tPBB.Text = "," Or tPBB.Text = "-" Then tPBB.Text = 0
If tDenda.Text = "" Or tDenda.Text = "." Or tDenda.Text = "," Or tDenda.Text = "-" Then tDenda.Text = 0
tPBB.Text = Format(tPBB.Text, "#,#0")
vOP.ListItems.Item(1).ListSubItems(16).Text = Format(tDenda.Text, "#,#0")
vOP.ListItems.Item(1).ListSubItems(17).Text = Format(tPBB.Text, "#,#0")
End Sub

Private Sub tSPPT_GotFocus()
On Error Resume Next
tSPPT.SelStart = 0
tSPPT.SelLength = Len(tSPPT.Text)
tSPPT.SetFocus
tSPPT.Alignment = 0

End Sub

Private Sub tSPPT_LostFocus()
On Error Resume Next
tSPPT.Alignment = 1
End Sub

Private Sub xUP_DownClick()
On Error Resume Next
If tNo.Text > 1 Then
    tNo.Text = tNo.Text - 1
End If
If hTunggal.Value = 1 Then cUbah
End Sub

Private Sub xUP_UpClick()
On Error Resume Next
tNo.Text = tNo.Text + 1
If hTunggal.Value = 1 Then cUbah
End Sub
Sub cUbah()
On Error Resume Next
d_YAR = "select * From PEMBAYARAN_SPPT where KD_PROPINSI + '.' + KD_DATI2  +'.' + KD_KECAMATAN +'.' + KD_KELURAHAN +'.' + KD_BLOK +'-' +NO_URUT +'.' +KD_JNS_OP= '" & aNOP.Text & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "' and PEMBAYARAN_SPPT_KE='" & tNo.Text & "'"
openDB (d_YAR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then
    tPBB.Text = rPajak!JML_SPPT_YG_DIBAYAR
Else
    tPBB.Text = 0
End If

End Sub
