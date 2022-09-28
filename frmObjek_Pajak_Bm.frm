VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmObjek_Pajak_Bm 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Data Objek Pajak Bumi"
   ClientHeight    =   8190
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   6390
   Icon            =   "frmObjek_Pajak_Bm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   120
      Picture         =   "frmObjek_Pajak_Bm.frx":1CCA
      ScaleHeight     =   330
      ScaleWidth      =   6180
      TabIndex        =   59
      Top             =   90
      Width           =   6180
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entri Data SPOP..."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   60
         TabIndex        =   60
         Top             =   75
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   270
         Left            =   5835
         Picture         =   "frmObjek_Pajak_Bm.frx":6332
         Stretch         =   -1  'True
         Top             =   30
         Width           =   300
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      TabIndex        =   64
      Top             =   855
      Width           =   6180
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
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
         Height          =   360
         Index           =   0
         Left            =   270
         TabIndex        =   65
         Top             =   -15
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.CommandButton cmdNOP1 
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
         Height          =   345
         Left            =   5655
         Picture         =   "frmObjek_Pajak_Bm.frx":8AD4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   90
         Width           =   375
      End
      Begin MSMask.MaskEdBox mBUMI 
         Height          =   330
         Left            =   1140
         TabIndex        =   3
         Top             =   105
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648384
         MaxLength       =   24
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "NOP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   66
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      TabIndex        =   51
      Top             =   1245
      Width           =   6180
      Begin VB.TextBox txtPajak 
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
         Index           =   1
         Left            =   1155
         TabIndex        =   5
         Top             =   195
         Width           =   2910
      End
      Begin VB.ComboBox cboNOP 
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
         Index           =   5
         Left            =   4695
         TabIndex        =   6
         Top             =   195
         Width           =   1395
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor SPOP"
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
         TabIndex        =   77
         Top             =   255
         Width           =   915
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4170
         TabIndex        =   76
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   120
      TabIndex        =   34
      Top             =   1680
      Width           =   6180
      Begin MSComCtl2.UpDown NTurun 
         Height          =   330
         Left            =   4425
         TabIndex        =   84
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboNOP 
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
         Index           =   1
         Left            =   1140
         TabIndex        =   8
         Top             =   510
         Width           =   4950
      End
      Begin VB.ComboBox cboNOP 
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
         Index           =   0
         Left            =   1140
         TabIndex        =   7
         Top             =   165
         Width           =   4950
      End
      Begin VB.ComboBox cboNOP 
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
         Index           =   2
         Left            =   1140
         TabIndex        =   9
         Top             =   870
         Width           =   1170
      End
      Begin VB.ComboBox cboNOP 
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
         Index           =   3
         Left            =   3345
         Style           =   1  'Simple Combo
         TabIndex        =   10
         Top             =   855
         Width           =   1080
      End
      Begin VB.ComboBox cboNOP 
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
         Index           =   4
         ItemData        =   "frmObjek_Pajak_Bm.frx":A79E
         Left            =   5205
         List            =   "frmObjek_Pajak_Bm.frx":A7A0
         TabIndex        =   11
         Top             =   870
         Width           =   885
      End
      Begin VB.Image iJalan 
         Height          =   375
         Left            =   2310
         MouseIcon       =   "frmObjek_Pajak_Bm.frx":A7A2
         MousePointer    =   99  'Custom
         Picture         =   "frmObjek_Pajak_Bm.frx":AAAC
         Stretch         =   -1  'True
         ToolTipText     =   "Pilih Blok/Nama Jalan"
         Top             =   840
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
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
         Left            =   4740
         TabIndex        =   40
         Top             =   930
         Width           =   360
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Urut"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2640
         TabIndex        =   39
         Top             =   915
         Width           =   930
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Blok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   38
         Top             =   915
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Kel/Desa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   37
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label41 
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
         Height          =   180
         Left            =   150
         TabIndex        =   35
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   6180
      TabIndex        =   81
      Top             =   5580
      Width           =   6180
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "PETUGAS"
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
         Height          =   210
         Left            =   0
         TabIndex        =   82
         Top             =   15
         Width           =   6195
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   6180
      TabIndex        =   78
      Top             =   3435
      Width           =   6180
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "LOKASI"
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
         Height          =   210
         Left            =   0
         TabIndex        =   80
         Top             =   30
         Width           =   6195
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Petugas"
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
         Left            =   1365
         TabIndex        =   79
         Top             =   30
         Width           =   1125
      End
   End
   Begin VB.Frame Frame6 
      Height          =   2190
      Left            =   6405
      TabIndex        =   67
      Top             =   105
      Width           =   6255
      Begin VB.TextBox tBumi 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   1260
         TabIndex        =   71
         Top             =   510
         Width           =   2265
      End
      Begin VB.TextBox tBumi 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   3540
         TabIndex        =   70
         Top             =   510
         Width           =   2505
      End
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   465
         TabIndex        =   69
         Top             =   510
         Width           =   780
      End
      Begin VB.CommandButton cmdHit 
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
         Height          =   315
         Left            =   60
         Picture         =   "frmObjek_Pajak_Bm.frx":C776
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   525
         Width           =   345
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   74
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total NJOP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1470
         TabIndex        =   73
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PBB Terutang"
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
         Left            =   4230
         TabIndex        =   72
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Height          =   660
         Left            =   420
         TabIndex        =   75
         Top             =   210
         Width           =   5670
      End
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   120
      TabIndex        =   45
      Top             =   3585
      Width           =   6180
      Begin VB.ComboBox cboJalan 
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
         Height          =   330
         Left            =   1170
         TabIndex        =   15
         Top             =   150
         Width           =   4905
      End
      Begin VB.TextBox tBumi 
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
         Index           =   9
         Left            =   2700
         MaxLength       =   2
         TabIndex        =   18
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox tBumi 
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
         Index           =   8
         Left            =   1170
         MaxLength       =   2
         TabIndex        =   17
         Top             =   855
         Width           =   1170
      End
      Begin VB.TextBox tBumi 
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
         Index           =   10
         Left            =   4785
         MaxLength       =   5
         TabIndex        =   19
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox tBumi 
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
         Index           =   11
         Left            =   1170
         TabIndex        =   16
         Top             =   495
         Width           =   4875
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Jalan"
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
         TabIndex        =   50
         Top             =   195
         Width           =   1365
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "RW"
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
         TabIndex        =   49
         Top             =   930
         Width           =   390
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blok/Kav/ID"
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
         TabIndex        =   48
         Top             =   555
         Width           =   840
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Persil"
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
         Left            =   4060
         TabIndex        =   47
         Top             =   915
         Width           =   675
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "RT"
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
         Left            =   2390
         TabIndex        =   46
         Top             =   915
         Width           =   1320
      End
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   120
      TabIndex        =   41
      Top             =   2820
      Width           =   6180
      Begin VB.ComboBox cboStatus 
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
         ItemData        =   "frmObjek_Pajak_Bm.frx":D440
         Left            =   4470
         List            =   "frmObjek_Pajak_Bm.frx":D453
         TabIndex        =   14
         Top             =   225
         Width           =   1620
      End
      Begin VB.CommandButton cmdID 
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
         Height          =   315
         Left            =   3240
         Picture         =   "frmObjek_Pajak_Bm.frx":D491
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Width           =   345
      End
      Begin VB.TextBox tID1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         TabIndex        =   12
         Top             =   210
         Width           =   2100
      End
      Begin VB.Label LAlamat 
         Caption         =   "Alamat WP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   165
         TabIndex        =   44
         Top             =   1485
         Width           =   1395
      End
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "Status WP"
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
         Left            =   3660
         TabIndex        =   43
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "No. KTP/ID"
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
         Left            =   150
         TabIndex        =   42
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   135
      TabIndex        =   85
      Top             =   4635
      Width           =   6165
      Begin VB.TextBox tBumi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4650
         TabIndex        =   23
         Text            =   "0"
         Top             =   510
         Width           =   1410
      End
      Begin VB.ComboBox cboJenis 
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
         Left            =   1170
         TabIndex        =   20
         Top             =   180
         Width           =   4935
      End
      Begin VB.TextBox tBumi 
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
         Index           =   5
         Left            =   1170
         TabIndex        =   21
         Text            =   "1"
         Top             =   525
         Width           =   720
      End
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   525
         Width           =   780
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Objek"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   165
         TabIndex        =   89
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Kod 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode ZNT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1980
         TabIndex        =   88
         Top             =   555
         Width           =   1545
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Luas Tanah"
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
         Left            =   3750
         TabIndex        =   87
         Top             =   555
         Width           =   825
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jlh Bangunan"
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
         TabIndex        =   86
         Top             =   555
         Width           =   975
      End
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   120
      TabIndex        =   52
      Top             =   5715
      Width           =   6180
      Begin VB.TextBox tBumi 
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
         Index           =   25
         Left            =   3855
         TabIndex        =   29
         Top             =   930
         Width           =   2250
      End
      Begin VB.TextBox tBumi 
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
         Index           =   24
         Left            =   3855
         TabIndex        =   27
         Top             =   600
         Width           =   2250
      End
      Begin VB.TextBox tBumi 
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
         Index           =   23
         Left            =   3855
         TabIndex        =   25
         Top             =   270
         Width           =   2250
      End
      Begin MSComCtl2.DTPicker dtPajak 
         Height          =   315
         Index           =   1
         Left            =   1290
         TabIndex        =   26
         Top             =   585
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   98238465
         CurrentDate     =   41486
      End
      Begin MSComCtl2.DTPicker dtPajak 
         Height          =   315
         Index           =   2
         Left            =   1290
         TabIndex        =   28
         Top             =   915
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   98238465
         CurrentDate     =   41486
      End
      Begin MSComCtl2.DTPicker dtPajak 
         Height          =   315
         Index           =   0
         Left            =   1290
         TabIndex        =   24
         Top             =   255
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   98238465
         CurrentDate     =   41486
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Pendataan"
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
         Left            =   105
         TabIndex        =   58
         Top             =   285
         Width           =   1980
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "NIP Perekam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2670
         TabIndex        =   57
         Top             =   975
         Width           =   1215
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Perekaman"
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
         TabIndex        =   56
         Top             =   945
         Width           =   1050
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "NIP Pemeriksa"
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
         Left            =   2685
         TabIndex        =   55
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Pemeriksaan"
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
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "NIP Pendata"
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
         Left            =   2670
         TabIndex        =   53
         Top             =   300
         Width           =   1260
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   120
      TabIndex        =   83
      Top             =   7440
      Width           =   6180
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Keluar"
         Top             =   180
         Width           =   630
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2505
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Batal"
         Top             =   180
         Width           =   660
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1830
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Simpan"
         Top             =   180
         Width           =   690
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7755
      Left            =   105
      TabIndex        =   36
      Top             =   330
      Width           =   6210
      Begin VB.TextBox tBumi 
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
         Index           =   7
         Left            =   1320
         TabIndex        =   30
         Top             =   6840
         Width           =   4770
      End
      Begin VB.CheckBox chPajak 
         BackColor       =   &H80000002&
         Caption         =   "Pemutakhiran Data"
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
         Index           =   2
         Left            =   2115
         TabIndex        =   1
         Top             =   270
         Width           =   1695
      End
      Begin VB.CheckBox chPajak 
         BackColor       =   &H80000002&
         Caption         =   "Penghapusan Data"
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
         Index           =   3
         Left            =   4215
         TabIndex        =   2
         Top             =   270
         Width           =   1665
      End
      Begin VB.CheckBox chPajak 
         BackColor       =   &H80000002&
         Caption         =   "Perekaman Data"
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
         Index           =   1
         Left            =   180
         TabIndex        =   0
         Top             =   270
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   90
         Top             =   6885
         Width           =   1215
      End
   End
   Begin VB.Label LTarif 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nilai Kelas"
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
      Left            =   10680
      TabIndex        =   63
      Top             =   6270
      Width           =   705
   End
   Begin VB.Label LNilai 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nilai Kelas"
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
      Left            =   10575
      TabIndex        =   62
      Top             =   5835
      Width           =   705
   End
   Begin VB.Label LKelas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIR"
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
      Left            =   10575
      TabIndex        =   61
      Top             =   5610
      Width           =   270
   End
End
Attribute VB_Name = "frmObjek_Pajak_Bm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jumRek, K1, K2, PBBMin
Dim xNIR, xNilai_Kelas
Dim xxTarif, yTarif
Dim xKelas
Dim jTrans
Dim totChar
Dim xxCabang
Dim xxTotalBng, xxNJOPBng
Dim Ikut
Private Sub cmdBangunan_Click()
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
frmOP_Tanah.Show
End Sub

Private Sub cboJalan_Click()
On Error Resume Next
tBumi(6).Text = Left(Trim(cboJalan.Text), 2)
End Sub

Private Sub cboJalan_DropDown()
On Error Resume Next
CALL_Jalan
End Sub

Private Sub cboJalan_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub cboJalan_LostFocus()
On Error Resume Next
cboJalan.Text = Rep(cboJalan.Text)
For i = 0 To cboJalan.ListCount - 1
        If (UCase(cboJalan.List(i)) Like "*" + UCase(cboJalan.Text) + "*" = True) Then
            cboJalan.Text = cboJalan.List(i)
            cboJalan_Click
            Exit Sub
        End If
          If i = cboJalan.ListCount - 1 Then
            If UCase(cboJalan.List(i)) Like "*" + UCase(cboJalan.Text) + "*" = False Then
                cboJalan.Text = cboJalan.List(0)
                cboJalan_Click
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub cboJenis_Click()
On Error Resume Next
If Left(Trim(cboJenis.Text), 2) = "01" Or Left(Trim(cboJenis.Text), 2) = "04" Then
    tBumi(5).Enabled = True
    tBumi(5).Locked = False
    tBumi(5).BackColor = vbWhite
    tBumi(5).Text = 1
Else
    tBumi(5).Enabled = False
    tBumi(5).Locked = True
    tBumi(5).BackColor = vbButtonFace
    tBumi(5).Text = 0
End If
End Sub

Private Sub cboJenis_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

If InStr("0123456789-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub cboJenis_LostFocus()
On Error Resume Next
 cboJenis.Text = cboJenis.List(cboJenis.Text - 1)
    If cboJenis.Text <> cboJenis.List(0) And cboJenis.Text <> cboJenis.List(1) And cboJenis.Text <> cboJenis.List(2) And cboJenis.Text <> cboJenis.List(3) Then
        cboJenis.Text = cboJenis.List(0)
    End If
    cboJenis_Click
End Sub

Private Sub cboNOP_Change(Index As Integer)
On Error Resume Next

Select Case Index
    Case 0
                If cboNOP(1).Text = "" Then iJalan.Visible = False Else iJalan.Visible = True

    Case 1
                If cboNOP(1).Text = "" Then iJalan.Visible = False Else iJalan.Visible = True

    Case 2
        'CALL_Jalan
        tBumi(0).Text = "12.12." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)
        mBUMI.Text = "12.12." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)
    Case 3
        'tBumi(0).Text = K1 & "." & K2 & "." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)
        tBumi(0).Text = "12.12." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)
        mBUMI.Text = "12.12." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)


    Case 4
End Select
End Sub


Private Sub cboNOP_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
  If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
Select Case Index
    Case 0 To 3, 5
        If InStr("0123456789-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
         KeyAscii = 0
        End If
    Case 4
        If InStr("0789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
         KeyAscii = 0
        End If
End Select
End Sub

Private Sub cboNOP_LostFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 0, 1, 2, 3, 4
  For i = 0 To cboNOP(Index).ListCount - 1
        If (UCase(cboNOP(Index).List(i)) Like "*" + UCase(cboNOP(Index).Text) + "*" = True) Then
            cboNOP(Index).Text = cboNOP(Index).List(i)
            cboNOP_Click (Index)
            GoTo Keluar
        End If
        'cboPajak(Index).Text = cboPajak(Index).List(cboPajak(Index).Text - 1)
     'If cboPajak(Index).Text = "" Then cboPajak(Index).Text = cboPajak(Index).List(0)
          If i = cboNOP(Index).ListCount - 1 Then
            If UCase(cboNOP(Index).List(i)) Like "*" + UCase(cboNOP(Index).Text) + "*" = False Then
                cboNOP(Index).Text = cboNOP(Index).List(0)
                'cDPA_Click (Index)
                cboNOP_Click (Index)
                GoTo Keluar
            End If
        End If
    Next
Case 5
    'If cboNOP(5).Text = "" Then cboNOP(5).Text = cboNOP(5).List(0)
     For i = 0 To cboNOP(Index).ListCount - 1
        If (UCase(cboNOP(Index).List(i)) Like "*" + UCase(cboNOP(Index).Text) + "*" = True) Then
            cboNOP(Index).Text = cboNOP(Index).List(i)
            Exit Sub
        End If
          If i = cboNOP(Index).ListCount - 1 Then
            If UCase(cboNOP(Index).List(i)) Like "*" + UCase(cboNOP(Index).Text) + "*" = False Then
                cboNOP(Index).Text = cboNOP(Index).List(0)
                Exit Sub
            End If
        End If
    Next
End Select
Keluar:
Select Case Index
Case 1
    callJRec
    If Trim(cboNOP(2).Text) = "000" Then cboNOP(4).Text = cboNOP(4).List(1) Else cboNOP(4).Text = cboNOP(4).List(0)
    tBumi(0).Text = "12.12." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)
    mBUMI.Text = "12.12." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)
End Select
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

If InStr("0123456789-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub cboStatus_LostFocus()
On Error Resume Next
cboStatus.Text = cboStatus.List(cboStatus.Text - 1)
     If cboStatus.Text = "" Then cboStatus.Text = cboStatus.List(0)
End Sub

'Private Sub cboZNT_DropDown()
'cboZNT.Clear
'strZNT = "Select * From DAT_PETA_ZNT where KD_KECAMATAN='" & Left(cboNOP(0).Text, 3) & "' and KD_KELURAHAN='" & Left(cboNOP(1).Text, 3) & "' and KD_BLOK='" & Trim(cboNOP(2).Text) & "' order by KD_ZNT"
''strZNT = "Select * From DAT_NIR where KD_KECAMATAN='" & Left(cboNOP(0).Text, 3) & "' and KD_KELURAHAN='" & Left(cboNOP(1).Text, 3) & "' order by KD_ZNT"
'openDB (strZNT)
'If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'Do While Not rPajak.EOF
'    cboZNT.AddItem rPajak!KD_ZNT
'    rPajak.MoveNext
'Loop
'End Sub

Private Sub chPajak_Click(Index As Integer)
On Error GoTo Salah

Select Case Index
Case 1
    If chPajak(1).Value = 1 Then
        chPajak(2).Value = 0
        chPajak(3).Value = 0
        cmdSave.Caption = "&Save"
        'cmdSubjek.Enabled = False
        cmdNOP1.Enabled = False
        mBUMI.Enabled = False
        CEK = 1
        cboNOP(2).Text = ""
            cboNOP(3).Text = ""
    Else
        If chPajak(2).Value = 0 And chPajak(3).Value = 0 Then
            chPajak(1).Value = 1
        End If
    End If
Case 2
    If chPajak(2).Value = 1 Then
        chPajak(1).Value = 0
        chPajak(3).Value = 0
        cmdNOP1.Enabled = True 'cmdSubjek.Enabled = True
        cmdSave.Caption = "&Update"
        CEK = 2
        xTanya = MsgBox("Apa anda yakin melakukan pemutakhiran Objek Pajak Bumi?", vbQuestion + vbYesNo, "Pemutakhiran")
        If xTanya = vbYes Then
            cmdSave.Caption = "&Update"
            xID = 1
            cboNOP(2).Text = ""
            cboNOP(3).Text = ""
            mBUMI.Enabled = True
        Else
            chPajak(1).Value = 1
            chPajak(2).Value = 0
            mBUMI.Enabled = False
        End If
     Else
        If chPajak(1).Value = 0 And chPajak(3).Value = 0 Then
            chPajak(2).Value = 1
        End If
    End If
   
Case 3
    If chPajak(3).Value = 1 Then
    chPajak(1).Value = 0
    chPajak(2).Value = 0
    'cmdSubjek.Enabled = True
    cmdNOP1.Enabled = True
    cmdSave.Caption = "&Delete"
    CEK = 3
    xTanya = MsgBox("Apa anda yakin menghapus Objek Pajak Bumi?", vbQuestion + vbYesNo, "Penghapusan")
    If xTanya = vbYes Then
        cmdSave.Caption = "&Delete"
        xID = 1
        cboNOP(2).Text = ""
        cboNOP(3).Text = ""
        mBUMI.Enabled = True
    Else
        chPajak(1).Value = 1
        chPajak(3).Value = 0
        mBUMI.Enabled = False
    End If
   Else
        If chPajak(1).Value = 0 And chPajak(2).Value = 0 Then
            chPajak(3).Value = 1
        End If
   End If
    
End Select


'Select Case Index
'Case 1
'    If chPajak(1).Value = 1 Then
'        chPajak(2).Value = 0
'        chPajak(3).Value = 0
'    End If
'    jTrans = 1
'Case 2
'
'   If chPajak(2).Value = 1 Then
'    chPajak(1).Value = 0
'    chPajak(3).Value = 0
'    xTanya = MsgBox("Apa anda yakin menghapus OBJEK PAJAK?", vbQuestion + vbYesNo, "Penghapusan NOP")
'    If xTanya = vbYes Then
'        xID = 1
'        frmList_Objek.Show
'    Else
'        chPajak(1).Value = 1
'        chPajak(2).Value = 0
'    End If
'
'   End If
'   jTrans = 2
'Case 3
'    If chPajak(3).Value = 1 Then
'        chPajak(1).Value = 0
'        chPajak(2).Value = 0
'        xTanya = MsgBox("Apa anda yakin melakukan pemutakhiran OBJEK PAJAK?", vbQuestion + vbYesNo, "Penghapusan NOP")
'        If xTanya = vbYes Then
'            xID = 1
'            frmList_Objek.Show
'        Else
'            chPajak(1).Value = 1
'            chPajak(3).Value = 0
'        End If
'    End If
'    jTrans = 3
'End Select
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub chPajak_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
Aktif
xID = ""
zJalan = ""

End Sub

Private Sub cmdExit_Click()
On Error Resume Next
Unload Me
xID = ""
zJalan = ""
'NO_FORM = 0
End Sub

Private Sub cmdHit_Click()
On Error GoTo Salah
Dim xTarif As Single
'tBumi(2).Text = 0
callNIR
callTarif

hitKelas1 (xNIR)
tBumi(2).Text = xKelas
tBumi(3).Text = Format(tBumi(1).Text * xNilai_Kelas, "#,#0")
BatasTarif (tBumi(3).Text)
'    If tBumi(3).Text * 1 >= 500000000 Then
'        xTarif = tBumi(3).Text * 1 - 15000000
'        tBumi(4).Text = Format(xTarif * 0.00222, "#,#0")
'    Else
'        xTarif = 0.00222 * (tBumi(3).Text * 1 - 10000000)
yTarif = xxTarif * tBumi(3).Text
        If yTarif < PBBMin Then
            tBumi(4).Text = PBBMin
        Else
            'tBumi(4).Text = Format(xTarif, "#,#0.00")
            tBumi(4).Text = Format(yTarif, "#,#0.00")
        End If
'    End If
    
    LTarif.Caption = xxTarif
LKelas.Caption = "NIR: " & xNIR 'tampil di luar form/lebarin az form
LNilai.Caption = "Nilai Kelas :" & xNilai_Kelas
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub cmdID_Click()
On Error Resume Next
xID = 1
Ikut = 0
frmList_Subjek.Show
End Sub

Private Sub cmdNOP1_Click()
On Error Resume Next
'On Error GoTo Salah
Ikut = 1
J_Karakter
If Len(Trim(tBumi(0).Text)) - (totChar * 1) = 24 Then
'MsgBox Len(mBUMI.Text)
'If mBUMI.SelLength = 24 Then
       'StrQ = "Select * From QOBJEKPAJAK WHERE NOPQ='" & Trim(tBumi(0).Text) & "' order by NOPQ asc"
       'StrQ = "Select * From DAT_OBJEK_PAJAK WHERE KD_PROPINSI +'.'+ KD_DATI2  +'.'+ KD_KECAMATAN  +'.'+ KD_KELURAHAN  +'.'+ KD_BLOK  +'-'+  NO_URUT  +'.'+  KD_JNS_OP ='" & Trim(mBUMI.Text) & "' "
       
       
       StrQ = "Select * From QOBJEKPAJAK WHERE NOPQ='" & Trim(mBUMI.Text) & "' ORDER BY NOPQ ASC"
        openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        MsgBox " Nomor Objek Pajak (NOP) Belum Terdaftar...", vbCritical + vbOKOnly, "Tidak Ketemu"
        Exit Sub
    End If
    Do While Not rPajak.EOF
        'IF LEFT(CBONOP(0).Text ,3)=
        If rPajak!KD_JNS_OP = 0 Then
            cboNOP(4).Text = cboNOP(4).List(0)
        ElseIf rPajak!KD_JNS_OP = 7 Then
            cboNOP(4).Text = cboNOP(4).List(1)
        ElseIf rPajak!KD_JNS_OP = 8 Then
            cboNOP(4).Text = cboNOP(4).List(2)
        ElseIf rPajak!KD_JNS_OP = 9 Then
            cboNOP(4).Text = cboNOP(4).List(3)
        End If
        cboNOP(0).Text = rPajak!KD_KECAMATAN & "-" & rPajak!NM_KECAMATAN
        cboNOP(1).Text = rPajak!KD_KELURAHAN & "-" & rPajak!NM_KELURAHAN
        cboNOP(2).Text = rPajak!KD_BLOK
        cboNOP(3).Text = rPajak![NO_URUT]
        cboNOP(4).Text = rPajak![KD_JNS_OP]
    'tBumi(0).Text = vBumi.SelectedItem.ListSubItems(2).Text 'NOP
        cboJenis.Text = cboJenis.List((rPajak!JNS_BUMI * 1) - 1) 'Jenis tanah
        tBumi(1).Text = rPajak!TOTAL_LUAS_BUMI 'Luas tanah
        txtPajak(1).Text = rPajak!NO_FORMULIR_SPOP 'Formulir/Dokumen
        
        'tID1.Text = rPajak!SUBJEK_PAJAK_ID 'ID Subjek Pajak
        If IsNull(rPajak!SUBJEK_PAJAK_ID) = True Or rPajak!SUBJEK_PAJAK_ID = "" Then
            tID1.Text = ""
        Else
            tID1.Text = rPajak!SUBJEK_PAJAK_ID
        End If
        cboStatus.Text = cboStatus.List((rPajak!KD_STATUS_WP * 1) - 1) 'Status Kepemilikan
        tBumi(6).Text = rPajak!KD_ZNT 'ZNT
        If IsNull(rPajak!JALAN_OP) = True Or rPajak!JALAN_OP = "" Then
             cboJalan.Text = "-"
        Else
             cboJalan.Text = rPajak!KD_ZNT & "-" & rPajak!JALAN_OP 'Nama Jalan
        End If
        
        If IsNull(rPajak!RW_OP) = True Or rPajak!RW_OP = "" Then
             tBumi(8).Text = "00"
        Else
         tBumi(8).Text = rPajak!RW_OP 'RW
         End If
        If IsNull(rPajak!RT_OP) = True Or rPajak!RT_OP = "" Then
             tBumi(9).Text = "00"
        Else
             tBumi(9).Text = rPajak!RT_OP 'RW
         End If
         If IsNull(rPajak!NO_PERSIL) = True Or rPajak!NO_PERSIL = "" Then
             tBumi(10).Text = "00"
        Else
             tBumi(10).Text = rPajak!NO_PERSIL
         End If
         If IsNull(rPajak!BLOK_KAV_NO_OP) = True Or rPajak!BLOK_KAV_NO_OP = "" Then
             tBumi(11).Text = "00"
        Else
             tBumi(11).Text = rPajak!BLOK_KAV_NO_OP
         End If
        
        dtPajak(0).Value = Format(rPajak!TGL_PENDATAAN_OP, "dd/mm/yyyy") 'Tanggal Pendataan
        dtPajak(1).Value = Format(rPajak!TGL_PEMERIKSAAN_OP, "dd/mm/yyyy") 'Tanggal Pemeriksaan
        dtPajak(2).Value = Format(rPajak!TGL_PEREKAMAN_OP, "dd/mm/yyyy") 'Tanggal Perekaman
        
        If IsNull(rPajak!NIP_PENDATA) = True Or rPajak!NIP_PENDATA = "" Then rPajak!NIP_PENDATA = "-"
        tBumi(23).Text = rPajak!NIP_PENDATA 'NIP Pendata
        If IsNull(rPajak!NIP_PEMERIKSA_OP) = True Or rPajak!NIP_PEMERIKSA_OP = "" Then rPajak!NIP_PEMERIKSA_OP = "-"
        tBumi(24).Text = rPajak!NIP_PEMERIKSA_OP 'NIP Pemeriksa
        If IsNull(rPajak!NIP_PEREKAM_OP) = True Or rPajak!NIP_PEREKAM_OP = "" Then rPajak!NIP_PEREKAM_OP = "-"
        tBumi(25).Text = rPajak!NIP_PEREKAM_OP 'NIP Perekam
        
        
       ' txtPajak(2).Text = " NAMA" & vbTab & ": " & rPajak!Nm_wp & vbCrLf & " LOKASI" & vbTab & ": " & rPajak!JALAN_OP & " Blok: " & rPajak!KD_BLOK & ", RT/RW: " & rPajak!RT_OP & "/" & rPajak!RW_OP & ", " & rPajak!NM_KELURAHAN & ", KEC. " & rPajak!NM_KECAMATAN 'NAMA dan Alamat
       rPajak.MoveNext
    Loop
    If chPajak(2).Value = 1 Or chPajak(3).Value = 1 Then tempLog1
Else
xID = 1
frmList_Objek.Show
End If
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'Salah:
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cmdNOP2_Click()
frmNOP.Show
End Sub

Private Sub cmdNOP3_Click()
frmNOP.Show
End Sub

Private Sub cmdSave_Click()
On Error GoTo Salah
xID = ""
zJalan = ""

If chPajak(1).Value = 1 Then
    CEK = 1
ElseIf chPajak(2).Value = 1 Then
    CEK = 2
ElseIf chPajak(3).Value = 1 Then
    CEK = 3
Else
    CEK = 1
End If
    

Select Case CEK
Case 1
    CTANYA = MsgBox("Apa Anda Yakin Menyimpan Data Objek Pajak Bumi?", vbQuestion + vbYesNo, "Simpan")
    If CTANYA = vbYes Then
        CALL_OPERASI (1)
        'Aktif
    End If
Case 2
    CTANYA = MsgBox("Apa Anda Yakin Mengupdate Data Objek Pajak Bumi?", vbQuestion + vbYesNo, "Update")
    If CTANYA = vbYes Then
        CALL_OPERASI (2)
'        Aktif
    End If
Case 3
    CTANYA = MsgBox("Apa Anda Yakin Menghapus Data Objek Pajak Bumi?", vbQuestion + vbYesNo, "Hapus")
    If CTANYA = vbYes Then
        CALL_OPERASI (3)
        'Aktif
    End If
    
End Select
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description


End Sub

Sub J_Karakter()
On Error Resume Next
 Dim jmlText, jmlChar, i As Integer
    
    jmlChar = 0 ' nol kan variabel JmlChar supaya hasil perhitungan valid
    jmlText = Len(tBumi(0).Text) 'Menghitung jumlah karakter didalam Text2 dan
                              'dan isikan hasilnya ke variabel JmlText
    
    For i = 0 To jmlText 'Melakukan proses perulangan sebanyak jumlah character dalam text2
'        tBumi(0).SetFocus
        tBumi(0).SelStart = i
        tBumi(0).SelLength = 1
        If tBumi(0).SelText = "_" Then 'Bandingkan Text2 dengan karakter yang dicari dari Text1
            jmlChar = jmlChar + 1 'Lakukan penambahan saat ditemukan karakter yang sesuai
        End If
    Next
    totChar = jmlChar
    'MsgBox "You have " & jmlChar & " Character " ' & tBumi(0).Text
    'Command1.SetFocus 'Mengembalikan posisi focus ke Command1"
End Sub

Private Sub dtPajak_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
Screen.MousePointer = vbHourglass
Dim C(100)
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
If xID = "" And zJalan = "" Then
For i = 0 To 2
        dtPajak(i).Value = Format(Now, "dd/mm/yyyy")
    Next
cboJenis.Clear
'cboJenis.Text = "01-Tanah dan Bangunan"
cboJenis.AddItem "01-TANAH DAN BANGUNAN"
cboJenis.AddItem "02-KAVLING SIAP BANGUN"
cboJenis.AddItem "03-TANAH KOSONG"
cboJenis.AddItem "04-FASILITAS UMUM"
cboJenis.Text = cboJenis.List(0)
'For i = 0 To 4
'    cboNOP(i).Clear
'Next
'cboNOP(3).Text = "0001"
cboNOP(4).Clear
'cboNOP(4).Text = "0-Peta"
cboNOP(4).AddItem "0" '-Peta"
cboNOP(4).AddItem "7" '-No Peta"
'cboNOP(4).AddItem "8" '-PB Peta"
'cboNOP(4).AddItem "9" '-PB No Peta"
cboNOP(4).Text = cboNOP(4).List(0)

    callKec
    Aktif
    cmdNOP1.Enabled = False
    mBUMI.Enabled = False
    tBumi(1).Text = 0
End If

'cboZNT.Clear
cboNOP(5).Clear
cboNOP(5).Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    cboNOP(5).AddItem i
Next
If Trim(cboNOP(0).Text) = "" Or Trim(cboNOP(1).Text) = "" Then
    iJalan.Visible = False
Else
    iJalan.Visible = True
End If
'For i = 0 To 99
'    cboZNT.AddItem Format(i, "00")
'Next
'C(1) = "A": C(2) = "B": C(3) = "C": C(4) = "D": C(5) = "E": C(6) = "F": C(7) = "G": C(8) = "H": C(9) = "I": C(10) = "J": C(11) = "K": C(12) = "L": C(13) = "M"
'C(14) = "N": C(15) = "O": C(16) = "P": C(17) = "Q": C(18) = "R": C(19) = "S": C(20) = "T": C(21) = "U": C(22) = "V": C(23) = "W": C(24) = "X": C(25) = "Y": C(26) = "Z"
'For J = 1 To 26 '676
'    For K = 1 To 26
'        cboZNT.AddItem C(J) & C(K)
'    Next
'Next
'    cboZNT.AddItem "B" & C(J)
'    cboZNT.AddItem "C" & C(J)
'    cboZNT.AddItem "D" & C(J)
'    cboZNT.AddItem "E" & C(J)
'    cboZNT.AddItem "F" & C(J)
'    cboZNT.AddItem "G" & C(J)
'    cboZNT.AddItem "H" & C(J)
'    cboZNT.AddItem "I" & C(J)
'    cboZNT.AddItem "J" & C(J)
'    cboZNT.AddItem "K" & C(J)
'    cboZNT.AddItem "L" & C(J)
'    cboZNT.AddItem "M" & C(J)
'    cboZNT.AddItem "N" & C(J)
'    cboZNT.AddItem "O" & C(J)
'    cboZNT.AddItem "P" & C(J)
'    cboZNT.AddItem "Q" & C(J)
'    cboZNT.AddItem "R" & C(J)
'    cboZNT.AddItem "S" & C(J)
'    cboZNT.AddItem "T" & C(J)
'    cboZNT.AddItem "U" & C(J)
'    cboZNT.AddItem "V" & C(J)
'    cboZNT.AddItem "W" & C(J)
'    cboZNT.AddItem "X" & C(J)
'    cboZNT.AddItem "Y" & C(J)
'    cboZNT.AddItem "Z" & C(J)
'Next
'MsgBox cboZNT.ListCount
If Ikut = 1 And (chPajak(2).Value = 1 Or chPajak(3).Value = 1) Then tempLog1
Screen.MousePointer = vbDefault
End Sub






Private Sub cboNOP_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
        cboNOP(1).Clear: cboNOP(2).Clear: cboNOP(3).Clear
        callKel
        If cboNOP(1).Text = "" Then iJalan.Visible = False Else iJalan.Visible = True
    Case 1
        If cboNOP(1).Text = "" Then iJalan.Visible = False Else iJalan.Visible = True
         If zJalan <> 1 Then
            callBlok
            CALL_Jalan
            cboNOP(3).Clear
        Else
            callJRec
        End If
         
    Case 2
        cboNOP(3).Clear
        callJRec
        callKab
        If Trim(cboNOP(2).Text) = "000" Then cboNOP(4).Text = cboNOP(4).List(1) Else cboNOP(4).Text = cboNOP(4).List(0)
        tBumi(0).Text = "12.12." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)
        mBUMI.Text = "12.12." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)
        'CALL_Jalan1
        If cboJalan.ListCount = 0 Then
            MsgBox "Nama Jalan Lokasi di BLOK [" & cboNOP(2).Text & "] Belum Terdaftar! " & _
            vbCrLf & "Silahkan Registrasi Nama Jalan atau Pilih BLOK Lain...", vbCritical + vbOKOnly, "TIDAK KETEMU"
        End If
        
    Case 4
        callKab
        tBumi(0).Text = "12.12." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)
        mBUMI.Text = "12.12." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cboNOP(2).Text, 3) & "-" & Format(cboNOP(3).Text, "0000") & "." & Left(cboNOP(4).Text, 1)
End Select
End Sub

Private Sub cboNOP_DropDown(Index As Integer)
Select Case Index
    Case 1
        'callKel
    Case 2
        'callBlok
'
        
End Select

End Sub





Sub callKec()
On Error GoTo Salah
cboNOP(0).Clear: cboNOP(1).Clear: cboNOP(2).Clear ': cboZNT.Clear
strKEC = "Select * From REF_KECAMATAN order by KD_KECAMATAN"
openDB (strKEC)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
cboNOP(0).AddItem rPajak!KD_KECAMATAN & "-" & rPajak!NM_KECAMATAN
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub callKel()
On Error GoTo Salah
cboNOP(1).Clear: cboNOP(2).Clear ': cboZNT.Clear
strKEL = "Select * From REF_KELURAHAN where KD_KECAMATAN='" & Left(cboNOP(0).Text, 3) & "' order by KD_KELURAHAN"
openDB (strKEL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
cboNOP(1).AddItem rPajak!KD_KELURAHAN & "-" & rPajak!NM_KELURAHAN
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub callBlok()
On Error GoTo Salah
cboNOP(2).Clear ': cboZNT.Clear
'strBLOK = "Select * From DAT_PETA_BLOK where KD_KECAMATAN='" & Left(cboNOP(0).Text, 3) & "' and KD_KELURAHAN='" & Left(cboNOP(1).Text, 3) & "' order by KD_BLOK"
strBLOK = "Select KD_BLOK ,KD_KECAMATAN,KD_KELURAHAN From JALAN where KD_KECAMATAN='" & Left(Trim(cboNOP(0).Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(cboNOP(1).Text), 3) & "' GROUP BY KD_KECAMATAN, KD_KELURAHAN, KD_BLOK ORDER BY KD_BLOK ASC"
openDB (strBLOK)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
cboNOP(2).AddItem rPajak!KD_BLOK
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub callJRec()
On Error GoTo Salah
Dim nURUT
StrJR = "Select * From DAT_OP_BUMI where KD_KECAMATAN='" & Left(cboNOP(0).Text, 3) & "' and KD_KELURAHAN='" & Left(cboNOP(1).Text, 3) & "' and KD_BLOK='" & Trim(cboNOP(2).Text) & "' order by NO_URUT, KD_ZNT"
openDB (StrJR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
jumRek = 0
Do While Not rPajak.EOF
    jumRek = jumRek + 1
    nURUT = rPajak!NO_URUT * 1
rPajak.MoveNext
Loop
'cboNOP(3).Text = Format(jumRek + 1, "0000")
If chPajak(2).Value = 1 Or chPajak(3).Value = 1 Then
    cboNOP(3).Text = Format(nURUT, "0000")
Else
    cboNOP(3).Text = Format(nURUT + 1, "0000")
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub


Sub callKab()
On Error GoTo Salah
strKab = "Select * From REF_DATI2"
openDB (strKab)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
K1 = rPajak!KD_PROPINSI
K2 = rPajak!KD_DATI2
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub callNIR()
On Error GoTo Salah
strKab = "Select * From DAT_NIR where KD_ZNT = '" & tBumi(6).Text & "' and THN_NIR_ZNT='" & Trim(cboNOP(5).Text) - 1 & "' and KD_KECAMATAN='" & Left(cboNOP(0).Text, 3) & "' and KD_KELURAHAN='" & Left(cboNOP(1).Text, 3) & "' order by THN_NIR_ZNT,KD_ZNT asc"
openDB (strKab)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    tBumi(2).Text = Format(Trim(rPajak!NIR), "#,#0")
        xNIR = Format(Trim(rPajak!NIR), "#,#0")
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub callTarif()
On Error GoTo Salah
strTarif = "Select * From TARIF"
openDB (strTarif)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'Do While Not rPajak.EOF
'    If tBumi(2).Text * 1 >= rPajak!NJOP_MIN And tBumi(2).Text * 1 <= rPajak!NJOP_MAX Then
'        tBumi(3).Text = Format(tBumi(2).Text * 1 * rPajak!NILAI_TARIF * 0.2, "#,#0")
'    End If
'rPajak.MoveNext
'Loop
    
strTarif = "Select * From PBB_MINIMAL Where THN_PBB_MINIMAL>='" & cboNOP(5).Text & "'"
openDB (strTarif)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
rPajak.Find "THN_PBB_MINIMAL='" & cboNOP(5).Text & "'"
If Not rPajak.EOF Then
    PBBMin = rPajak!NILAI_PBB_MINIMAL
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
xxID = 0
End Sub
Sub hitKelas1(BUMI As Single)
On Error GoTo Salah
Dim xxTAwal
StrQ = "Select THN_AWAL_KLS_TANAH From KELAS_TANAH  order by THN_AWAL_KLS_TANAH asc"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        xxTAwal = rPajak!THN_AWAL_KLS_TANAH
    rPajak.MoveNext
    Loop
    
    '====
    StrQ = "Select * From KELAS_TANAH WHERE THN_AWAL_KLS_TANAH='" & xxTAwal & "'"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        If BUMI >= rPajak!NILAI_MIN_TANAH * 1 And BUMI <= rPajak!NILAI_MAX_TANAH * 1 Then
            xKelas = rPajak!KD_KLS_TANAH
            xNilai_Kelas = rPajak!NILAI_PER_M2_TANAH ' * 1000
'            xTot = rPajak!KD_KLS_TANAH * Format(rPajak!NILAI_PER_M2_TANAH * 1000, "#,#0")
            Exit Sub
        End If
    rPajak.MoveNext
    Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
            
    
End Sub
Sub BatasTarif(xNJOP As Single)
On Error GoTo Salah
StrQ = "Select * From TARIF"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        If xNJOP * 1 >= rPajak!NJOP_MIN * 1 And xNJOP * 1 <= rPajak!NJOP_MAX * 1 Then
            xxTarif = rPajak!NILAI_TARIF * 0.01
            Exit Sub
        End If
    rPajak.MoveNext
    Loop
            
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_Jalan()
On Error GoTo Salah
cboJalan.Clear
StrQ = "Select * From JALAN where KD_KECAMATAN='" & Left(Trim(cboNOP(0).Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(cboNOP(1).Text), 3) & "' ORDER BY KD_BLOK,KD_ZNT ASC"
openDB (StrQ)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        'If rPajak!NM_JLN_STANDARD = "" Or IsNull(rPajak!NM_JLN_STANDARD) = True Then
         '   cboJalan.AddItem rPajak!NM_JLN_SEMENTARA
        'Else
         '   cboJalan.AddItem rPajak!NM_JLN_STANDARD
        'End If
        cboJalan.AddItem rPajak!KD_ZNT & "-" & rPajak!NM_JLN
    rPajak.MoveNext
    Loop
            
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_Jalan1()
On Error GoTo Salah
cboJalan.Clear
StrQ = "Select * From JALAN where KD_KECAMATAN='" & Left(Trim(cboNOP(0).Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(cboNOP(1).Text), 3) & "' AND KD_BLOK='" & Trim(cboNOP(2).Text) & "'  ORDER BY KD_BLOK,KD_ZNT ASC"
openDB (StrQ)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        'If rPajak!NM_JLN_STANDARD = "" Or IsNull(rPajak!NM_JLN_STANDARD) = True Then
         '   cboJalan.AddItem rPajak!NM_JLN_SEMENTARA
        'Else
         '   cboJalan.AddItem rPajak!NM_JLN_STANDARD
        'End If
        cboJalan.AddItem rPajak!KD_ZNT & "-" & rPajak!NM_JLN
    rPajak.MoveNext
    Loop
            
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub iJalan_Click()
zJalan = 1
Ikut = 0
frmJalan.Show
End Sub

Private Sub Image1_Click()
Unload Me
xID = ""
'NO_FORM = 0
End Sub

Sub CALL_OPERASI(CEK1)
'On Error GoTo Salah
Dim xxKec, xxKel, xxBlok, xxUrut, xxJenis, xxNOP, NOPQ
xxKec = Left(Trim(cboNOP(0).Text), 3)
xxKel = Left(Trim(cboNOP(1).Text), 3)
xxBlok = Left(Trim(cboNOP(2).Text), 3)
xxUrut = Trim(cboNOP(3).Text)
xxJenis = Left(Trim(cboNOP(4).Text), 1)
xxNOP = "12.12." & xxKec & "." & xxKel & "." & xxBlok & "-" & xxUrut & "." & xxJenis
cmdHit_Click
'xSTR = "Select * From DAT_OP_BUMI" 'where KD_KECAMATAN='" & xxKec & "' AND KD_KELURAHAN='" & xxKel & "' AND KD_BLOK='" & xxBlok & "' AND NO_URUT='" & xxUrut & "' AND KD_JNS_OP='" & xxJenis & "'"
xSTR = "Select * From DAT_OP_BUMI where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & xxNOP & "'"
openDB (xSTR)

If rPajak.RecordCount > 0 Then rPajak.MoveFirst

For Each Control In Me
    If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
        If Control.Text = "" Then
            MsgBox "Masih ada Data Kosong, Tolong dilengkapi...." & Control.Name, vbCritical, "Tetnong"
            GoTo Keluar

        End If

    End If
Next
QJalan = Mid(Trim(cboJalan.Text), 4, Len(Trim(cboJalan.Text)))
Select Case CEK1
Case 1
    If Not rPajak.EOF Then 'Jika Ditemukan
        MsgBox "Data Sudah Ada...", vbCritical, "Error"
       ' MsgBox "12.12." + rPajak!KD_KECAMATAN + "." + rPajak!KD_KELURAHAN + "." + rPajak!KD_BLOK + "-" + rPajak!NO_URUT + "." + rPajak!KD_JNS_OP
          'MsgBox xxNOP
        Exit Sub
    End If
    xSTR = "select * from DAT_SUBJEK_PAJAK WHERE SUBJEK_PAJAK_ID='" & tID1.Text & "' ORDER BY SUBJEK_PAJAK_ID ASC"
    openDB (xSTR)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        MsgBox "Subjek Pajak Belum Terdaftar, Silahkan Lihat di tabel...", vbCritical, "Tetnong"
        Exit Sub
    End If
    
'    rPajak.AddNew
'    rPajak!KD_PROPINSI = "12"
'    rPajak!KD_DATI2 = "12"
'    rPajak!KD_KECAMATAN = xxKec 'Left(Trim(cboNOP(0).Text), 3)
'    rPajak!KD_KELURAHAN = xxKel 'Left(Trim(cboNOP(1).Text), 3)
'    rPajak!KD_BLOK = xxBlok 'Left(Trim(cboNOP(2).Text), 3)
'    rPajak!NO_URUT = xxUrut 'Trim(cboNOP(3).Text)
'    rPajak!KD_JNS_OP = xxJenis 'Left(Trim(cboNOP(4).Text), 1)
'    rPajak!NO_BUMI = 1
'    rPajak!KD_ZNT = tBumi(6).Text
'    rPajak!LUAS_BUMI = tBumi(1).Text
'    rPajak!JNS_BUMI = Left(Trim(cboJENIS.Text), 2)*1
'    rPajak!NILAI_SISTEM_BUMI = tBumi(3).Text
'    rPajak.Update

'----Simpan ke Tabel DAT_OP_BUMI
    'iSQL = "insert into DAT_OP_BUMI(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,NO_BUMI,KD_ZNT,LUAS_BUMI,JNS_BUMI, NILAI_SISTEM_BUMI)" & _
            "Values('12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '1', '" & tBumi(6).Text & "', '" & Round(tBumi(1).Text, 0) & "', '" & Left(Trim(cboJenis.Text), 2) * 1 & "', '" & Round(tBumi(3).Text, 0) & "')"
    'openDB (iSQL)
'----Simpan ke Tabel Objek Pajak
    'QJalan = Mid(Trim(cboJalan.Text), 4, Len(Trim(cboJalan.Text)))
    'iSQL3 = "insert into DAT_OBJEK_PAJAK(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,SUBJEK_PAJAK_ID,NO_FORMULIR_SPOP,NO_PERSIL,JALAN_OP, BLOK_KAV_NO_OP,RW_OP,RT_OP,KD_STATUS_WP,TOTAL_LUAS_BUMI,NJOP_BUMI,JNS_TRANSAKSI_OP,TGL_PENDATAAN_OP,NIP_PENDATA,TGL_PEMERIKSAAN_OP,NIP_PEMERIKSA_OP,TGL_PEREKAMAN_OP,NIP_PEREKAM_OP,KD_STATUS_CABANG,TOTAL_LUAS_BNG,NJOP_BNG,STATUS_PETA_OP)" & _
            "Values('12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '" & tID1.Text & "', '" & txtPajak(1).Text & "', '" & tBumi(10).Text & "', '" & QJalan & "', '" & tBumi(11).Text & "', '" & tBumi(8).Text & "', '" & tBumi(9).Text & "', '" & Left(Trim(cboStatus.Text), 1) * 1 & "', '" & tBumi(1).Text & "', '" & tBumi(3).Text * 1000 & "', '1', '" & dtPajak(0).Value & "', '" & tBumi(23).Text & "', '" & dtPajak(1).Value & "', '" & tBumi(24).Text & "', '" & dtPajak(2).Value & "', '" & tBumi(25).Text & "','0',0,0,'1')"
    'openDB (iSQL3)
    C_STR = "INSERT_BUMI '12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '1', '" & tBumi(6).Text & "', '" & Round(tBumi(1).Text, 0) & "', '" & Left(Trim(cboJenis.Text), 2) * 1 & "', '" & Round(tBumi(3).Text, 0) & "','" & txtPajak(1).Text & "', '0'," & _
            "'12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '" & tID1.Text & "', '" & txtPajak(1).Text & "', '" & tBumi(10).Text & "', '" & QJalan & "', '" & tBumi(11).Text & "', '" & tBumi(8).Text & "', '" & tBumi(9).Text & "', '" & Left(Trim(cboStatus.Text), 1) * 1 & "', '" & tBumi(1).Text & "', '" & tBumi(3).Text * 1000 & "', '1', '" & Format(dtPajak(0).Value, "yyyy-mm-dd") & "', '" & tBumi(23).Text & "', '" & Format(dtPajak(1).Value, "yyyy-mm-dd") & "', '" & tBumi(24).Text & "', '" & Format(dtPajak(2).Value, "yyyy-mm-dd") & "', '" & tBumi(25).Text & "','0',0,0,'1'"
    openDB (C_STR)
    If Left(cboJenis.Text, 2) = "01" Or (Left(cboJenis.Text, 2) = "04" And tBumi(5).Text * 1 > 0) Then
    'MsgBox "Lanjutkan Mengisi Data Bangunan...", vbOKOnly + vbExclamation, "Objek Pajak"
    'Unload Me
       ' NO_FORM = txtPajak(1).Text
        byPass = Left(Trim(cboJenis.Text), 2)
        frmObjek_Pajak_Bg.aNOP.Text = tBumi(0).Text
        'frmObjek_Pajak_Bg.txtPajak(1).Text = tBumi(0).Text
        'frmObjek_Pajak_Bg.cboPajak(2).Text = cboNOP(5).Text
        'frmObjek_Pajak_Bg.cTPajak.Text = cboNOP(5).Text
        BYPASS1 = tBumi(0).Text
        BYPASS2 = cboNOP(5).Text
        BYPASS3 = tID1.Text
        bypass4 = 1
        xID = ""
        frmObjek_Pajak_Bg.Show
        
    End If
    Aktif
Case 2
    If rPajak.EOF Then 'Jika Ditemukan
        MsgBox "Data tidak berhasil di edit, karena belum ada di database...", vbCritical, "Error"
        Exit Sub
    End If
    xSTR = "select * from DAT_SUBJEK_PAJAK WHERE SUBJEK_PAJAK_ID='" & tID1.Text & "' ORDER BY SUBJEK_PAJAK_ID ASC"
    openDB (xSTR)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        MsgBox "Subjek Pajak Belum Terdaftar, Silahkan Lihat di tabel...", vbCritical, "Tetnong"
        Exit Sub
    End If

''    upObjek = "Select * From DAT_OBJEK_PAJAK where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & xxNOP & "'"
''    openDB (upObjek)
''    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
''    Do While Not rPajak.EOF
''        'If (Left(Trim(cboJenis.Text), 2) * 1 = 1 Or Left(Trim(cboJenis.Text), 2) * 1 = 4) And rPajak!NJOP_BNG <> 0 Then
''        If rPajak!TOTAL_LUAS_BNG <> 0 Or rPajak!NJOP_BNG <> 0 Then
''           MsgBox "Anda tidak dapat mengubah jenis objek/tanah..." & _
''            vbCrLf & "Hapus Objek Bangunan terlebih dahulu...", vbCritical, "Tetnong..."
''           Exit Sub
''        End If
''    rPajak.MoveNext
''    Loop
'
    upBumi = "Select * From DAT_OP_BUMI where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & xxNOP & "'"
    openDB (upBumi)

    'cmdHit_Click
    'Update data DAT_OP_BUMI
    rPajak!KD_PROPINSI = "12"
    rPajak!KD_DATI2 = "12"
    rPajak!KD_KECAMATAN = xxKec 'Left(Trim(cboNOP(0).Text), 3)
    rPajak!KD_KELURAHAN = xxKel 'Left(Trim(cboNOP(1).Text), 3)
    rPajak!KD_BLOK = xxBlok 'Left(Trim(cboNOP(2).Text), 3)
    rPajak!NO_URUT = xxUrut 'Trim(cboNOP(3).Text)
    rPajak!KD_JNS_OP = xxJenis 'Left(Trim(cboNOP(4).Text), 1)
    rPajak!NO_BUMI = 1
    rPajak!KD_ZNT = tBumi(6).Text
    rPajak!LUAS_BUMI = tBumi(1).Text
    rPajak!JNS_BUMI = Left(Trim(cboJenis.Text), 2) * 1
    rPajak!NILAI_SISTEM_BUMI = tBumi(3).Text
    rPajak!SUBJEK_PAJAK_ID = tID1.Text
    rPajak.Update


    'Update data DAT_OBJEK_PAJAK
    upSTR = "Select * From DAT_OBJEK_PAJAK where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & xxNOP & "'"
    openDB (upSTR)
'    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    Do While Not rPajak.EOF
'        If (Left(Trim(cboJenis.Text), 2) * 1 = 1 Or Left(Trim(cboJenis.Text), 2) * 1 = 4) And rPajak!NJOP <> 0 Then
'
'        End If
'    rPajak.MoveNext
'    Loop
    rPajak!KD_PROPINSI = "12"
    rPajak!KD_DATI2 = "12"
    rPajak!KD_KECAMATAN = xxKec 'Left(Trim(cboNOP(0).Text), 3)
    rPajak!KD_KELURAHAN = xxKel 'Left(Trim(cboNOP(1).Text), 3)
    rPajak!KD_BLOK = xxBlok 'Left(Trim(cboNOP(2).Text), 3)
    rPajak!NO_URUT = xxUrut 'Trim(cboNOP(3).Text)
    rPajak!KD_JNS_OP = xxJenis 'Left(Trim(cboNOP(4).Text), 1)
    rPajak!SUBJEK_PAJAK_ID = tID1.Text 'ID Subjek Pajak
   rPajak!NO_FORMULIR_SPOP = txtPajak(1).Text 'Formulir/Dokumen
   rPajak!NO_PERSIL = tBumi(10).Text 'Persil
   rPajak!JALAN_OP = Mid(Trim(cboJalan.Text), 4, Len(Trim(cboJalan.Text))) 'Nama Jalan
   rPajak!BLOK_KAV_NO_OP = tBumi(11).Text 'Blok/Kav
   rPajak!RW_OP = tBumi(8).Text 'RW
   rPajak!RT_OP = tBumi(9).Text 'RT
   'rPajak!KD_STATUS_CABANG = 0
   rPajak!KD_STATUS_WP = Left(Trim(cboStatus.Text), 1)
   rPajak!TOTAL_LUAS_BUMI = tBumi(1).Text
   'rPajak!TOTAL_LUAS_BNG = 0
   rPajak!NJOP_BUMI = tBumi(3).Text * 1000
   'rPajak!NJOP_BNG= 0
   'rPajak!STATUS_PETA_OP= "1"
   rPajak!JNS_TRANSAKSI_OP = CEK1
    rPajak!TGL_PENDATAAN_OP = Format(dtPajak(0).Value, "dd/mm/yyyy") 'Tanggal Pendataan
   rPajak!TGL_PEMERIKSAAN_OP = Format(dtPajak(1).Value, "dd/mm/yyyy") 'Tanggal Pemeriksaan
   rPajak!TGL_PEREKAMAN_OP = Format(dtPajak(2).Value, "dd/mm/yyyy") 'Tanggal Perekaman
   rPajak!NIP_PENDATA = tBumi(23).Text 'NIP Pendata
   rPajak!NIP_PEMERIKSA_OP = tBumi(24).Text 'NIP Pemeriksa
   rPajak!NIP_PEREKAM_OP = tBumi(25).Text 'NIP Perekam
    rPajak.Update
    
    'UP_STR = "UPDATE_BUMI '12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '1', '" & tBumi(6).Text & "', '" & Round(tBumi(1).Text * 1, 0) & "', '" & Left(Trim(cboJenis.Text), 2) * 1 & "', '" & Round(tBumi(3).Text, 0) & "','" & txtPajak(1).Text & "', '0','" & xxNOP & "'," & _
            "'12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '" & tID1.Text & "', '" & txtPajak(1).Text & "', '" & tBumi(10).Text & "', '" & QJalan & "', '" & tBumi(11).Text & "', '" & tBumi(8).Text & "', '" & tBumi(9).Text & "', '" & Left(Trim(cboStatus.Text), 1) * 1 & "', '" & tBumi(1).Text & "', '" & tBumi(3).Text * 1000 & "', '1', '" & dtPajak(0).Value & "', '" & tBumi(23).Text & "', '" & dtPajak(1).Value & "', '" & tBumi(24).Text & "', '" & dtPajak(2).Value & "', '" & tBumi(25).Text & "','0',0,0,'1','" & xxNOP & "'"
            'UP_STR = "UPDATE_BUMI '12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '1', '" & tBumi(6).Text & "', '" & (tBumi(1).Text * 1) & "', '" & Left(Trim(cboJenis.Text), 2) * 1 & "', '" & (tBumi(3).Text * 1) & "','" & txtPajak(1).Text & "', '0','" & xxNOP & "'," & _
            "'12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '" & tID1.Text & "', '" & txtPajak(1).Text & "', '" & tBumi(10).Text & "', '" & QJalan & "', '" & tBumi(11).Text & "', '" & tBumi(8).Text & "', '" & tBumi(9).Text & "', '" & Left(Trim(cboStatus.Text), 1) * 1 & "', '" & tBumi(1).Text & "', '" & tBumi(3).Text * 1000 & "', '1', '" & Format(dtPajak(0).Value, "yyyy-mm-dd") & "', '" & tBumi(23).Text & "', '" & Format(dtPajak(1).Value, "yyyy-mm-dd") & "', '" & tBumi(24).Text & "', '" & Format(dtPajak(2).Value, "yyyy-mm-dd") & "', '" & tBumi(25).Text & "',0,0,0,1,'" & xxNOP & "'"
    'openDB (UP_STR)
    
    Log1
    If Left(cboJenis.Text, 2) = "01" Or (Left(cboJenis.Text, 2) = "04" And tBumi(5).Text * 1 > 0) Then
        CTANYA = MsgBox("Anda telah berhasil mengedit seluruh data objek pajak bumi" & _
                vbCrLf & "Apa Anda Ingin Mengedit Data Bangunan ?", vbInformation + vbYesNo, "Updated...")
        If CTANYA = vbYes Then
            'NO_FORM = txtPajak(1).Text
            byPass = Left(Trim(cboJenis.Text), 2)
            frmObjek_Pajak_Bg.aNOP.Text = tBumi(0).Text
            BYPASS1 = tBumi(0).Text
            BYPASS2 = cboNOP(5).Text
            BYPASS3 = tID1.Text
            bypass4 = 2
            xID = ""
            frmObjek_Pajak_Bg.Show
        End If
    End If
    Aktif
Case 3
    If rPajak.EOF Then 'Jika Ditemukan
        MsgBox "Tidak ada data yang akan terhapus...", vbCritical, "Error"
        Exit Sub
    End If
'    If rPajak!JNS_BUMI = 1 Or rPajak!JNS_BUMI = 4 Then
'        Tanya = MsgBox("Objek Pajak Ini terdiri dari BUMI dan BANGUNAN," & _
'                vbCrLf & "Apa anda yakin menghapus kedua objek tersebut?,,,", vbYesNo + vbInformation, "Deleted...")
'        If Tanya = vbNo Then
'            Exit Sub
'        End If
'    End If
   upObjek = "Select * From DAT_OBJEK_PAJAK where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & xxNOP & "'"
    openDB (upObjek)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        'If (Left(Trim(cboJenis.Text), 2) * 1 = 1 Or Left(Trim(cboJenis.Text), 2) * 1 = 4) And rPajak!NJOP_BNG <> 0 Then
        If rPajak!TOTAL_LUAS_BNG <> 0 Or rPajak!NJOP_BNG <> 0 Then
           MsgBox "Objek Pajak Bumi dan Bangunan tidak dapat dihapus sekaligus..." & _
            vbCrLf & "Hapus Objek Bangunan terlebih dahulu...", vbCritical, "Tetnong..."
           Exit Sub
        End If
    rPajak.MoveNext
    Loop
    Log2
'    upBumi = "Select * From DAT_OP_BUMI where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & xxNOP & "'"
'    openDB (upBumi)
'        Do While Not rPajak.EOF
'            rPajak.Delete adAffectCurrent
'            rPajak.Update
'            rPajak.MoveNext
'        Loop
'        delSTR = "Select * From DAT_OBJEK_PAJAK where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & xxNOP & "'"
'        openDB (delSTR)
'        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'        Do While Not rPajak.EOF
'            rPajak.Delete adAffectCurrent
'            rPajak.Update
'            rPajak.MoveNext
'        Loop
        DEL_STR = "DELETE_BUMI '" & xxNOP & "','" & xxNOP & "'"
        openDB (DEL_STR)
        Aktif
End Select
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'Salah:
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'MsgBox Err.Number & ": " & Err.Description
Keluar:


End Sub

Sub Aktif()
On Error Resume Next
'tID(0).Text = "-"
    txtPajak(1).Text = ""
    cboNOP(0).Text = ""
    cboNOP(1).Text = ""
    cboNOP(2).Text = ""
    cboNOP(3).Text = ""
    cboNOP(4).Text = cboNOP(4).List(0)
    cboNOP(5).Text = cboNOP(5).List(0)
    tBumi(0).Text = ""
    For i = 0 To 2
        dtPajak(i).Value = Format(Now, "dd/mm/yyyy")
    Next
   mBUMI.Mask = ""
   mBUMI.Text = ""
    tID1.Text = ""
    cboStatus.Text = ""
'    cboJalan.Clear
    tBumi(8).Text = "00"
    tBumi(9).Text = "00"
    tBumi(10).Text = "00"
    tBumi(11).Text = "00"
    tBumi(1).Text = 0
    tBumi(6).Text = ""
    tBumi(5).Text = 1
    tBumi(23).Text = "-"
    tBumi(24).Text = "-"
    tBumi(25).Text = "-"
    tBumi(7).Text = "-"
    cboJenis.Text = cboJenis.List(0)
 '   cboStatus.Text = cboStatus.List(0)
    
End Sub


Private Sub Image2_Click()

End Sub

Private Sub mBUMI_Change()
On Error Resume Next
tBumi(0).Text = mBUMI.Text
End Sub

Private Sub mBUMI_GotFocus()
On Error Resume Next
mBUMI.Mask = "12.12.###.###.###-####.#"
End Sub

Private Sub mBUMI_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub mBUMI_LostFocus()
On Error Resume Next
tBumi(0).Text = mBUMI.Text
End Sub

Private Sub NTurun_DownClick()
On Error Resume Next
cboNOP(3).Text = Format(cboNOP(3).Text - 1, "0000")
If chPajak(3).Value = 1 Or chPajak(2).Value = 1 Then
    cmdNOP1_Click
End If
End Sub

Private Sub NTurun_UpClick()
On Error Resume Next
cboNOP(3).Text = Format(cboNOP(3).Text + 1, "0000")
If chPajak(3).Value = 1 Or chPajak(2).Value = 1 Then
    cmdNOP1_Click
End If

End Sub

Private Sub tBumi_GotFocus(Index As Integer)
On Error Resume Next
Select Case Index
    Case 1, 5, 8 To 11, 23, 24, 25
        Call c_blok(tBumi(Index))
End Select
End Sub

Private Sub tBumi_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case Index
Case 0
    If KeyCode = 13 Then
        If chPajak(1).Value = 0 Then
            cmdNOP1_Click
        End If
    End If
End Select
End Sub

Private Sub tBumi_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
  If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
Select Case Index
Case 0, 1, 5, 8 To 10
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select
End Sub

Private Sub tBumi_LostFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 1, 23, 24, 25
    tBumi(Index).Text = Rep(tBumi(Index).Text)
    Call c_Kosong(tBumi(Index), 0)
Case 5
    Call c_Kosong(tBumi(Index), 1)
Case 7
    tBumi(7).Text = Rep(tBumi(7).Text)
Case 8 To 11
    tBumi(11).Text = Rep(tBumi(11).Text)
    Call c_Kosong(tBumi(Index), "00")
End Select
End Sub

Private Sub tID1_GotFocus()
On Error Resume Next
Call c_blok(tID1)
End Sub

Private Sub tID1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    call_Subjek
End If
End Sub


Private Sub txNOP_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
 If InStr("0123456789,", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tID1_LostFocus()
tID1.Text = Rep(tID1.Text)
End Sub

Private Sub txtPajak_GotFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 1
    Call c_blok(txtPajak(Index))
End Select
End Sub

Private Sub txtPajak_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
  If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
Select Case Index
Case 1
  
    If InStr("0123456789-.,", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select
End Sub
Sub call_Subjek()
On Error GoTo Salah
xSTR = "select * from DAT_SUBJEK_PAJAK WHERE SUBJEK_PAJAK_ID='" & tID1.Text & "' ORDER BY SUBJEK_PAJAK_ID ASC"
    openDB (xSTR)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        MsgBox "Subjek Pajak Belum Terdaftar, Silahkan Lihat di tabel...", vbCritical, "Tetnong"
        Exit Sub
    End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub c_blok(nControl As TextBox)
On Error Resume Next
nControl.SelStart = 0
nControl.SelLength = Len(nControl.Text)
nControl.SetFocus
nControl.Alignment = 0
End Sub
Sub c_Kosong(nControl As TextBox, xNilai)
On Error Resume Next
If nControl.Text = "" Or nControl.Text = "-" Or nControl.Text = "." Then
    nControl.Text = xNilai
End If
nControl.Alignment = 1
End Sub

Private Sub txtPajak_LostFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 1
    Call c_Kosong(txtPajak(Index), 0)
End Select
End Sub
Sub tempLog1()
On Error Resume Next
cmdHit_Click
'panggil PBB terutang dari tabel SPPT
xSkg = Format(Now, "yyyy")
xxSPPT = "select * from SPPT where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(mBUMI.Text) & "' AND THN_PAJAK_SPPT='" & xSkg * 1 - 1 & "'"
openDB (xxSPPT)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    ccPBB = rPajak!PBB_YG_HARUS_DIBAYAR_SPPT
    ccNJOPTKP = rPajak!NJOPTKP_SPPT
    ccKelas1 = rPajak!KD_KLS_TANAH
    ccKelas2 = rPajak!KD_KLS_BNG
    If rPajak!LUAS_BNG_SPPT = "" Or IsNull(rPajak!LUAS_BNG_SPPT) = True Or rPajak!LUAS_BNG_SPPT = Null Then
        cTotalBNG = 0
    Else
        cTotalBNG = rPajak!LUAS_BNG_SPPT
    End If
    If rPajak!NJOP_BNG_SPPT = "" Or IsNull(rPajak!NJOP_BNG_SPPT) = True Or rPajak!NJOP_BNG_SPPT = Null Then
        cNJOPBng = 0
    Else
'        cNJOPBng = rPajak!tt!njop_bng_sppt
         cNJOPBng = rPajak!NJOP_BNG_SPPT
    End If
    
    rPajak.MoveNext
Loop
'panggil nama sebelum dirubah
dNama = "select * from DAT_SUBJEK_PAJAK WHERE SUBJEK_PAJAK_ID='" & Trim(tID1.Text) & "' ORDER  BY NM_WP ASC"
openDB (dNama)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    ccNama = rPajak!Nm_wp
    rPajak.MoveNext
Loop


upSTR = "Delete tempLogUtama"
    openDB (upSTR)

upSTR = "Select * From tempLogUtama"
    openDB (upSTR)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    rPajak.AddNew
    rPajak!KD_PROPINSI = "12"
    rPajak!KD_DATI2 = "12"
    rPajak!KD_KECAMATAN = Left(Trim(cboNOP(0).Text), 3)
    rPajak!KD_KELURAHAN = Left(Trim(cboNOP(1).Text), 3)
    rPajak!KD_BLOK = Left(Trim(cboNOP(2).Text), 3)
    rPajak!NO_URUT = Trim(cboNOP(3).Text)
    rPajak!KD_JNS_OP = Left(Trim(cboNOP(4).Text), 1)
    rPajak!SUBJEK_PAJAK_ID = tID1.Text 'ID Subjek Pajak
   rPajak!NO_FORMULIR_SPOP = txtPajak(1).Text 'Formulir/Dokumen
   rPajak!NO_PERSIL = tBumi(10).Text 'Persil
   rPajak!JALAN_OP = Mid(Trim(cboJalan.Text), 4, Len(Trim(cboJalan.Text))) & ", " & Mid(Trim(cboNOP(1).Text), 5, Len(cboNOP(1).Text)) & " KEC. " & Mid(Trim(cboNOP(0).Text), 5, Len(cboNOP(0).Text)) 'Nama Jalan
   rPajak!BLOK_KAV_NO_OP = tBumi(11).Text 'Blok/Kav
   rPajak!RW_OP = tBumi(8).Text 'RW
   rPajak!RT_OP = tBumi(9).Text 'RT
   rPajak!NO_BNG = 0
   rPajak!KD_STATUS_WP = Left(Trim(cboStatus.Text), 1)
   rPajak!TOTAL_LUAS_BUMI = tBumi(1).Text
    rPajak!NJOP_BUMI = tBumi(3).Text * 1000
    rPajak!PBB_Terutang = ccPBB
    rPajak!NJOPTKP = ccNJOPTKP
    rPajak!KD_STATUS_CABANG = 0
        If cTotalBNG = Empty Then
            rPajak!TOTAL_LUAS_BNG = 0
        Else
            rPajak!TOTAL_LUAS_BNG = cTotalBNG
        End If
        If cNJOPBng = Empty Then
            rPajak!NJOP_BNG = 0
        Else
            rPajak!NJOP_BNG = cNJOPBng
        End If
    

   
   rPajak!STATUS_PETA_OP = "1"
   If chPajak(2).Value = 1 Then
        rPajak!JNS_TRANSAKSI_OP = 2
   Else
        rPajak!JNS_TRANSAKSI_OP = 3
   End If
    rPajak!TGL_PENDATAAN_OP = Format(dtPajak(0).Value, "dd/mm/yyyy") 'Tanggal Pendataan
   rPajak!TGL_PEMERIKSAAN_OP = Format(dtPajak(1).Value, "dd/mm/yyyy") 'Tanggal Pemeriksaan
   rPajak!TGL_PEREKAMAN_OP = Format(dtPajak(2).Value, "dd/mm/yyyy") 'Tanggal Perekaman
   rPajak!NIP_PENDATA = tBumi(23).Text 'NIP Pendata
   rPajak!NIP_PEMERIKSA_OP = tBumi(24).Text 'NIP Pemeriksa
   rPajak!NIP_PEREKAM_OP = tBumi(25).Text 'NIP Perekam
   rPajak!KD_KLS_TANAH = ccKelas1
   rPajak!KD_KLS_BNG = ccKelas2
   rPajak!Nm_wp = ccNama
    rPajak.Update
End Sub
'Log untuk perubahan data bumi
Sub Log1()
On Error Resume Next
TANYA = MsgBox("Simpan di Log Data?", vbQuestion + vbYesNo, "Log...")
If TANYA = vbNo Then
    Exit Sub
End If

ccSTR = "Select  * from templogutama"
openDB (ccSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
cNOP = "12.12." & rPajak!KD_KECAMATAN & "." & rPajak!KD_KELURAHAN & "." & rPajak!KD_BLOK & "-" & rPajak!NO_URUT & "." & rPajak!KD_JNS_OP
    cKec = rPajak!KD_KECAMATAN
    cKel = rPajak!KD_KELURAHAN
    cblok = rPajak!KD_BLOK
    cUrut = rPajak!NO_URUT
    cJenis = rPajak!KD_JNS_OP
    cID = rPajak!SUBJEK_PAJAK_ID
    cLokasi = rPajak!JALAN_OP
    cWP = rPajak!KD_STATUS_WP
    cTotal1 = rPajak!TOTAL_LUAS_BUMI
    cBumi = rPajak!NJOP_BUMI
    cNama = rPajak!Nm_wp
    cPBB = rPajak!PBB_Terutang
    cNJOPTKP = rPajak!NJOPTKP
    cKelas1 = rPajak!KD_KLS_TANAH
    cKelas2 = rPajak!KD_KLS_BNG
    cLuasB = rPajak!TOTAL_LUAS_BNG
    cNJOPB = rPajak!NJOP_BNG
rPajak.MoveNext
Loop
'panggil nama setelah dirubah
dNama = "select * from DAT_SUBJEK_PAJAK WHERE SUBJEK_PAJAK_ID='" & Trim(tID1.Text) & "' ORDER  BY NM_WP ASC"
openDB (dNama)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    ccNama1 = rPajak!Nm_wp
    rPajak.MoveNext
Loop
    upSTR = "Select * From LogUtama where NOP1='" & Trim(cNOP) & "' order by NOP1 asc"
    openDB (upSTR)
   If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        rPajak.AddNew
        rPajak!TOTAL_LUAS_BNG = cLuasB
        rPajak!NJOP_BNG = cNJOPB
        rPajak!NO_BNG = 0
        rPajak!TOTAL_LUAS_BNG1 = cLuasB
        rPajak!NJOP_BNG1 = cNJOPB
        rPajak!NJOPTKP = cNJOPTKP
   'rPajak!NJOPTKP1 = 0

        'rPajak!TOTAL_LUAS_BNG1 = 0 ' xxTotalBng
    'rPajak!NJOP_BNG1 = 0 'xxNJOPBng
    
    End If
'    rPajak!KD_PROPINSI = "12"
'    rPajak!KD_DATI2 = "12"
'    rPajak!KD_KECAMATAN = Left(Trim(cboNOP(0).Text), 3)
'    rPajak!KD_KELURAHAN = Left(Trim(cboNOP(1).Text), 3)
'    rPajak!KD_BLOK = Left(Trim(cboNOP(2).Text), 3)
'    rPajak!NO_URUT = Trim(cboNOP(3).Text)
'    rPajak!KD_JNS_OP = Left(Trim(cboNOP(4).Text), 1)
    'Data Lama
     rPajak!NOP = "12.12." & cKec & "." & cKel & "." & cblok & "-" & cUrut & "." & cJenis
     rPajak!SUBJEK_PAJAK_ID = cID
     rPajak!Nm_wp = cNama
     rPajak!Lokasi = cLokasi
     rPajak!KD_STATUS_WP = cWP
     rPajak!TOTAL_LUAS_BUMI = cTotal1
     rPajak!NJOP_BUMI = cBumi
    'Data Baru
     vKec = Left(Trim(cboNOP(0).Text), 3)
    vKel = Left(Trim(cboNOP(1).Text), 3)
    vblok = Left(Trim(cboNOP(2).Text), 3)
    vUrut = Trim(cboNOP(3).Text)
    vJenis = Left(Trim(cboNOP(4).Text), 1)
    ccNOP = "12.12." & vKec & "." & vKel & "." & vblok & "-" & vUrut & "." & vJenis
    rPajak!NOP1 = ccNOP
    rPajak!subjek_pajak_id1 = tID1.Text 'ID Subjek Pajak
   rPajak!lokasi1 = Mid(Trim(cboJalan.Text), 4, Len(Trim(cboJalan.Text))) & ", " & Mid(Trim(cboNOP(1).Text), 5, Len(cboNOP(1).Text)) & " KEC. " & Mid(Trim(cboNOP(0).Text), 5, Len(cboNOP(0).Text)) 'Nama Jalan
   rPajak!kd_status_wp1 = Left(Trim(cboStatus.Text), 1)
   rPajak!TOTAL_LUAS_BUMI1 = tBumi(1).Text
   'MsgBox tBumi(1).Text * 1 & "--" & cTotal1 * 1
   If tBumi(1).Text * 1 <> cTotal1 * 1 Then rPajak!xFlag = "1" 'Else rPajak!xFlag = "0"
   rPajak!NJOP_BUMI1 = tBumi(3).Text * 1000
   If chPajak(2).Value = 1 Then
        rPajak!JNS_TRANSAKSI_OP = 2
   Else
        rPajak!JNS_TRANSAKSI_OP = 3
   End If
   'rPajak!TGL_PEREKAMAN_OP = Format(dtPajak(2).Value, "dd/mm/yyyy hh:mm:ss")  'Tanggal Perekaman
   rPajak!TGL_PEREKAMAN_OP = Format(Now, "dd/mm/yyyy hh:mm:ss")  'Tanggal Perekaman
   rPajak!NIP_PEREKAM_OP = tBumi(25).Text 'NIP Perekam
   'rPajak!TOTAL_LUAS_BNG = 0 ' xxTotalBng
   rPajak!NM_WP1 = ccNama1
   
   'rPajak!NJOP_BNG = 0 'xxNJOPBng
   'rPajak!NO_BNG = 0
   rPajak!PBB_Terutang = cPBB
   
   rPajak!PBB_TERUTANG1 = 0
   rPajak!KD_KLS_TANAH = tBumi(2).Text
   'rPajak!KD_KLS_BNG = cKelas2
   'rPajak!KD_KLS_BNG = cKelas2
   
  If rPajak!NOP <> rPajak!NOP1 Then c1 = "(Perubahan NOP) " Else c1 = ""
   If rPajak!Lokasi <> rPajak!lokasi1 Then c2 = "(Perubahan Lokasi WP) " Else c2 = ""
   If rPajak!Nm_wp <> rPajak!NM_WP1 Then c3 = "(Perubahan Nama WP) " Else c3 = ""
   If rPajak!TOTAL_LUAS_BUMI <> rPajak!TOTAL_LUAS_BUMI1 Then c4 = "(Perubahan Luas Bumi) " Else c4 = ""
   If rPajak!TOTAL_LUAS_BNG <> rPajak!TOTAL_LUAS_BNG1 Then c5 = "(Perubahan Luas Bangunan) " Else c5 = ""
   'If rPajak!NJOP_BUMI <> rPajak!NJOP_BUMI1 Then c6 = "(Perubahan NJOP Bumi) " Else c6 = ""
   'If rPajak!NJOP_BNG <> rPajak!NJOP_BNG1 Then c7 = "(Perubahan NJOP Bangunan) " Else c7 = ""
   If rPajak!NJOPTKP <> rPajak!NJOPTKP1 Then c8 = "(Perubahan NJOPTKP)" Else c8 = ""
    rPajak!KET = tBumi(7).Text & " " & c1 & c2 & c3 & c4 & c5 & c6 & c7 & c8
'   rajak!KD_KLS_TANAH = tBumi(2).Text
   
    rPajak.Update
   
End Sub

'Log untuk penghapusan data bumi


Sub Log2()
On Error Resume Next
TANYA = MsgBox("Simpan di Log Data?", vbQuestion + vbYesNo, "Log...")
If TANYA = vbNo Then
    Exit Sub
End If

ccSTR = "Select  * from templogutama"
openDB (ccSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    cKec = rPajak!KD_KECAMATAN
    cKel = rPajak!KD_KELURAHAN
    cblok = rPajak!KD_BLOK
    cUrut = rPajak!NO_URUT
    cJenis = rPajak!KD_JNS_OP
    cID = rPajak!SUBJEK_PAJAK_ID
    cLokasi = rPajak!JALAN_OP
    cWP = rPajak!KD_STATUS_WP
    cTotal1 = rPajak!TOTAL_LUAS_BUMI
    cBumi = rPajak!NJOP_BUMI
    cNama = rPajak!Nm_wp
rPajak.MoveNext
Loop



upSTR = "Select * From LogUtama where NOP1='" & Trim(cNOP) & "' order by NOP1 asc"
    openDB (upSTR)
   If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        rPajak.AddNew
        If cLuasB = Empty Then cLuasB = 0
        If cNJOPB = Empty Then cNJOPB = 0
        rPajak!TOTAL_LUAS_BNG = cLuasB
        rPajak!NJOP_BNG = cNJOPB
        rPajak!NO_BNG = 0
        rPajak!TOTAL_LUAS_BNG1 = cLuasB
        rPajak!NJOP_BNG1 = cNJOPB
        If cNJOPTKP = Empty Then cNJOPTKP = 0
        rPajak!NJOPTKP = cNJOPTKP
   
    End If
'
    'Data Lama
     rPajak!NOP = "12.12." & cKec & "." & cKel & "." & cblok & "-" & cUrut & "." & cJenis
     rPajak!SUBJEK_PAJAK_ID = cID
     rPajak!Nm_wp = cNama
     rPajak!Lokasi = cLokasi
     rPajak!KD_STATUS_WP = cWP
     rPajak!TOTAL_LUAS_BUMI = cTotal1
     rPajak!NJOP_BUMI = cBumi
    'Data Baru
     vKec = Left(Trim(cboNOP(0).Text), 3)
    vKel = Left(Trim(cboNOP(1).Text), 3)
    vblok = Left(Trim(cboNOP(2).Text), 3)
    vUrut = Trim(cboNOP(3).Text)
    vJenis = Left(Trim(cboNOP(4).Text), 1)
    ccNOP = "12.12." & vKec & "." & vKel & "." & vblok & "-" & vUrut & "." & vJenis
    rPajak!NOP1 = ccNOP
    rPajak!subjek_pajak_id1 = tID1.Text 'ID Subjek Pajak
   rPajak!lokasi1 = Mid(Trim(cboJalan.Text), 4, Len(Trim(cboJalan.Text))) & ", " & Mid(Trim(cboNOP(1).Text), 5, Len(cboNOP(1).Text)) & " KEC. " & Mid(Trim(cboNOP(0).Text), 5, Len(cboNOP(0).Text)) 'Nama Jalan
   rPajak!kd_status_wp1 = Left(Trim(cboStatus.Text), 1)
   rPajak!TOTAL_LUAS_BUMI1 = tBumi(1).Text
   
   rPajak!xFlag = "2"
   rPajak!NJOP_BUMI1 = tBumi(3).Text * 1000
   If chPajak(2).Value = 1 Then
        rPajak!JNS_TRANSAKSI_OP = 2
   Else
        rPajak!JNS_TRANSAKSI_OP = 3
   End If
   'rPajak!TGL_PEREKAMAN_OP = Format(dtPajak(2).Value, "dd/mm/yyyy hh:mm:ss")  'Tanggal Perekaman
   rPajak!TGL_PEREKAMAN_OP = Format(Now, "dd/mm/yyyy hh:mm:ss")  'Tanggal Perekaman
   rPajak!NIP_PEREKAM_OP = tBumi(25).Text 'NIP Perekam
   'rPajak!TOTAL_LUAS_BNG = 0 ' xxTotalBng
   rPajak!NM_WP1 = ccNama1
   
   'rPajak!NJOP_BNG = 0 'xxNJOPBng
   'rPajak!NO_BNG = 0
   rPajak!PBB_Terutang = cPBB
   
   rPajak!PBB_TERUTANG1 = 0
   'CEK NOP APAKAH SAMA
   
    rPajak!KET = tBumi(7).Text & " (Hapus Objek Pajak)"
   'rPajak!KD_KLS_BNG = cKelas2
    rPajak.Update
    
    strFlag = "Select * from logutama where NOP='" & mBUMI.Text & "' order by NOP asc"
   openDB (strFlag)
   If rPajak.RecordCount > 0 Then rPajak.MoveFirst
   Do While Not rPajak.EOF
    rPajak!xFlag = "2"
    rPajak.MoveNext
   Loop
   
End Sub
