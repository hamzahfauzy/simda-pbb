VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCetak_Massal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak SPPT Secara Massal"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16050
   ControlBox      =   0   'False
   Icon            =   "frmCetak_Massal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   16050
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   -30
      ScaleHeight     =   480
      ScaleWidth      =   6000
      TabIndex        =   24
      Top             =   -30
      Width           =   6000
      Begin VB.CheckBox cUrut 
         BackColor       =   &H80000002&
         Caption         =   "Cetak Berdasarkan NOP"
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
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   2580
      End
      Begin VB.CheckBox hTunggal 
         BackColor       =   &H80000002&
         Caption         =   "Cetak Tunggal"
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
         Left            =   120
         TabIndex        =   2
         Top             =   105
         Width           =   1875
      End
      Begin VB.CheckBox cRekam 
         BackColor       =   &H80000002&
         Caption         =   "Berdasarkan Tanggal Rekam Objek"
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
         Left            =   2895
         TabIndex        =   1
         Top             =   135
         Width           =   2730
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      Height          =   510
      Left            =   1890
      TabIndex        =   29
      Top             =   -75
      Width           =   3945
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
         Left            =   3555
         TabIndex        =   31
         Top             =   165
         Width           =   345
      End
      Begin MSMask.MaskEdBox aNOP 
         Height          =   315
         Left            =   510
         TabIndex        =   3
         Top             =   150
         Width           =   3045
         _ExtentX        =   5371
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
         Left            =   495
         TabIndex        =   30
         Top             =   150
         Width           =   3000
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
         Left            =   135
         TabIndex        =   32
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   6030
      TabIndex        =   36
      Top             =   855
      Width           =   5955
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Material"
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
         Index           =   15
         Left            =   4245
         TabIndex        =   53
         Top             =   1605
         Width           =   1575
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fasilitas III"
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
         Index           =   14
         Left            =   4245
         TabIndex        =   52
         Top             =   1365
         Width           =   1575
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fasilitas II"
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
         Index           =   13
         Left            =   4245
         TabIndex        =   51
         Top             =   1140
         Width           =   1575
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fasilitas I"
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
         Index           =   12
         Left            =   4245
         TabIndex        =   50
         Top             =   915
         Width           =   1575
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gedung Sekolah"
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
         Index           =   11
         Left            =   90
         TabIndex        =   49
         Top             =   3315
         Width           =   3495
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tangki Minyak"
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
         Index           =   10
         Left            =   90
         TabIndex        =   48
         Top             =   3090
         Width           =   3495
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kanopi Pompa Bensin, Daya Dukung dan Mezanin"
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
         Index           =   9
         Left            =   90
         TabIndex        =   47
         Top             =   2850
         Width           =   4365
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Apartemen"
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
         Index           =   8
         Left            =   90
         TabIndex        =   46
         Top             =   2625
         Width           =   3495
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bengkel/Gedung/Pertanian"
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
         Index           =   7
         Left            =   90
         TabIndex        =   45
         Top             =   2370
         Width           =   3495
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hotel/Wisma"
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
         Index           =   6
         Left            =   90
         TabIndex        =   44
         Top             =   2130
         Width           =   3495
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Olahraga/Rekreasi dan Bangunan Parkir"
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
         Index           =   5
         Left            =   90
         TabIndex        =   43
         Top             =   1875
         Width           =   3495
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rumah Sakit/Klinik"
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
         Index           =   4
         Left            =   90
         TabIndex        =   42
         Top             =   1635
         Width           =   3495
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pertokoan"
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
         Left            =   90
         TabIndex        =   41
         Top             =   1395
         Width           =   3495
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pabrik"
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
         Left            =   90
         TabIndex        =   40
         Top             =   1140
         Width           =   3495
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Perkantoran Swasta"
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
         Left            =   90
         TabIndex        =   39
         Top             =   885
         Width           =   3495
      End
      Begin VB.CheckBox chJPB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bangunan Standard"
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
         Left            =   90
         TabIndex        =   37
         Top             =   255
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(Perumahan, Kantor, Apotik, Pasar/Ruko,Restoran,Hotel,Wisma, Gedung Pemerintahan, Rumah Sakit/Klinik)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   315
         TabIndex        =   38
         Top             =   450
         Width           =   5250
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
      TabIndex        =   17
      Top             =   4590
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
      TabIndex        =   16
      Top             =   4590
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Cetak"
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
      TabIndex        =   15
      Top             =   4590
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   -60
      ScaleHeight     =   765
      ScaleWidth      =   6150
      TabIndex        =   25
      Top             =   4425
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
      Height          =   2970
      Left            =   -30
      TabIndex        =   18
      Top             =   345
      Width           =   5985
      Begin MSComCtl2.DTPicker dRekam2 
         Height          =   315
         Left            =   3720
         TabIndex        =   8
         Top             =   1200
         Width           =   1995
         _ExtentX        =   3519
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
         Format          =   97910785
         CurrentDate     =   41486
      End
      Begin MSComCtl2.DTPicker dRekam1 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   1200
         Width           =   1830
         _ExtentX        =   3228
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
         Format          =   97910785
         CurrentDate     =   41486
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
         TabIndex        =   4
         Top             =   195
         Width           =   1350
      End
      Begin VB.TextBox tTotal2 
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
         Left            =   3735
         TabIndex        =   33
         Text            =   "0"
         Top             =   1560
         Width           =   2010
      End
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
         Left            =   1515
         TabIndex        =   14
         Text            =   "0"
         Top             =   2565
         Width           =   2685
      End
      Begin VB.TextBox tTotal 
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
         Left            =   1515
         TabIndex        =   11
         Text            =   "0"
         Top             =   1560
         Width           =   1845
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
         TabIndex        =   6
         Top             =   855
         Width           =   4260
      End
      Begin VB.TextBox tSPPT 
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
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   1215
         Width           =   1845
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
         TabIndex        =   5
         Top             =   525
         Width           =   4260
      End
      Begin MSComCtl2.DTPicker dCetak 
         Height          =   300
         Left            =   1515
         TabIndex        =   13
         Top             =   2235
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
         Format          =   97910785
         CurrentDate     =   41486
      End
      Begin MSComCtl2.DTPicker dTerbit 
         Height          =   315
         Left            =   1515
         TabIndex        =   12
         Top             =   1905
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
         Format          =   97910785
         CurrentDate     =   41486
      End
      Begin VB.TextBox tSPPT2 
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
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   1230
         Width           =   1995
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "s.d"
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
         Left            =   3435
         TabIndex        =   54
         Top             =   1275
         Width           =   285
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
         Height          =   165
         Left            =   165
         TabIndex        =   22
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "NIP Pencetak"
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
         TabIndex        =   28
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Terbit"
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
         TabIndex        =   27
         Top             =   1920
         Width           =   1260
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Cetak"
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
         TabIndex        =   26
         Top             =   2265
         Width           =   1215
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
         TabIndex        =   23
         Top             =   1605
         Width           =   1305
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   1260
         Width           =   1305
      End
   End
   Begin MSComctlLib.ListView vOP 
      Height          =   3960
      Left            =   6150
      TabIndex        =   34
      Top             =   -15
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
      NumItems        =   52
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
         SubItemIndex    =   2
         Text            =   "NAMA KELURAHAN"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NAMA KECAMATAN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "SPPT1"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "JUMLAH1"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "SPPT2"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "JUMLAH2"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "TAHUN"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "JNS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "TPAJAK"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "SIKLUS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "KANWIL"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "KPBB"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "BANK1"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "BANK2"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "KD_TP"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "NAMA WP"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "ALAMAT WP"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   19
         Text            =   "KAV"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   20
         Text            =   "RW"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   21
         Text            =   "RT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "KELURAHAN"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   23
         Text            =   "KOTA"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "POS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "NPWP"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   26
         Text            =   "NOPERSIL"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   27
         Text            =   "K_TNH"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   28
         Text            =   "THN"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   29
         Text            =   "K_BNG"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   30
         Text            =   "THN"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "J_TEMPO"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Text            =   "L_BUMI"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   33
         Text            =   "L_BNG"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   34
         Text            =   "NJOP_BM"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "NJOP_BNG"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   36
         Text            =   "TOTAL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   37
         Text            =   "NJOPTKP"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   38
         Text            =   "NJKP"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   39
         Text            =   "BAYAR"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   40
         Text            =   "KURANG"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   41
         Text            =   "TOTAL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "ST_BYR"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   43
         Text            =   "STS_TAGIH"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "CETAK"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "T_TERBIT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "T_CETAK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "NIP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "PROSES"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "FLAG_NJOPTKP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "NOP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "J_Bumi"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView VOP1 
      Height          =   3960
      Left            =   6045
      TabIndex        =   35
      Top             =   3945
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
      NumItems        =   52
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
         SubItemIndex    =   2
         Text            =   "NAMA KELURAHAN"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NAMA KECAMATAN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "SPPT1"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "JUMLAH1"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "SPPT2"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "JUMLAH2"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "TAHUN"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "JNS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "TPAJAK"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "SIKLUS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "KANWIL"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "KPBB"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "BANK1"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "BANK2"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "KD_TP"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "NAMA WP"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "ALAMAT WP"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   19
         Text            =   "KAV"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   20
         Text            =   "RW"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   21
         Text            =   "RT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "KELURAHAN"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   23
         Text            =   "KOTA"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "POS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "NPWP"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   26
         Text            =   "NOPERSIL"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   27
         Text            =   "K_TNH"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   28
         Text            =   "THN"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   29
         Text            =   "K_BNG"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   30
         Text            =   "THN"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "J_TEMPO"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Text            =   "L_BUMI"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   33
         Text            =   "L_BNG"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   34
         Text            =   "NJOP_BM"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "NJOP_BNG"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   36
         Text            =   "TOTAL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   37
         Text            =   "NJOPTKP"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   38
         Text            =   "NJKP"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   39
         Text            =   "BAYAR"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   40
         Text            =   "KURANG"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   41
         Text            =   "TOTAL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "ST_BYR"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   43
         Text            =   "STS_TAGIH"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "CETAK"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "T_TERBIT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "T_CETAK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "NIP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "PROSES"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "FLAG_NJOPTKP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "NOP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "J_Bumi"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmCetak_Massal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xTT, xTB, QQ
Dim NMIN, cTarif
'Dim xMIN(2), xMAX(2)
Dim xTarif(2)
Dim cMin(2), cMax(2), cTKP(2)
Dim totChar
Private Sub cmdBangunan_Click()
On Error Resume Next
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
On Error Resume Next
frmOP_Tanah.Show
End Sub

Private Sub ccKel_Click()
On Error Resume Next
C_KEC = Left(Trim(ccKec.Text), 3)
    C_KEL = Left(Trim(ccKel.Text), 3)
    tSPPT2.Refresh
            tSPPT.Refresh
            C_STR = "SELECT * FROM QOBJEKPAJAK WHERE KD_KECAMATAN='" & C_KEC & "' AND KD_KELURAHAN='" & C_KEL & "' ORDER BY NO_URUT"
            openDB (C_STR)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            tSPPT.Text = rPajak!NO_URUT 'Mid(rPajak!NOPQ, 19, 4)
            If rPajak.RecordCount > 0 Then rPajak.MoveLast
            tSPPT2.Text = rPajak!NO_URUT 'Mid(rPajak!NOPQ, 19, 4)
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

Private Sub chJPB_Click(Index As Integer)
On Error GoTo Salah
Select Case Index
Case 0
    If chJPB(0).Value = 1 Then
        For i = 1 To 15
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "01"
Case 1
    If chJPB(1).Value = 1 Then
        For i = 2 To 15
            chJPB(i).Value = 0
        Next
            chJPB(0).Value = 0
    End If
    K_JPB = "02"
Case 2
    If chJPB(2).Value = 1 Then
        For i = 3 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 1
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "03"
Case 3
    If chJPB(3).Value = 1 Then
        For i = 4 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 2
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "04"
Case 4
    If chJPB(4).Value = 1 Then
        For i = 5 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 3
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "05"
Case 5
    If chJPB(5).Value = 1 Then
        For i = 6 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 4
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "12"
Case 6
    If chJPB(6).Value = 1 Then
        For i = 7 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 5
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "07"
Case 7
    If chJPB(7).Value = 1 Then
        For i = 8 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 6
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "08"
Case 8
    If chJPB(8).Value = 1 Then
        For i = 9 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 7
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "13"
Case 9
    If chJPB(9).Value = 1 Then
        For i = 10 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 8
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "14"
Case 10
    If chJPB(10).Value = 1 Then
        For i = 11 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 9
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "15"
Case 11
    If chJPB(11).Value = 1 Then
        For i = 12 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 10
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "16"
Case 12
    If chJPB(12).Value = 1 Then
        For i = 13 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 11
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "17"
Case 13
    If chJPB(13).Value = 1 Then
        For i = 14 To 15
            chJPB(i).Value = 0
        Next
        For i = 0 To 12
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "18"
Case 14
    If chJPB(14).Value = 1 Then
        chJPB(15).Value = 0
        For i = 0 To 13
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "19"
Case 15
    If chJPB(15).Value = 1 Then
        For i = 0 To 14
            chJPB(i).Value = 0
        Next
    End If
    K_JPB = "20"
End Select
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub cmdCear_Click()
On Error Resume Next
ccTahun.Text = ccTahun.List(0)
ccKec.Text = ""
ccKel.Text = ""
tSPPT.Text = 0
tTotal.Text = 0
dTerbit.Value = Format(Now, "dd/mm/yyyy")
dCetak.Value = Format(Now, "dd/mm/yyyy")
tNIP.Text = 0
hTunggal.Value = 0
Frame4.Visible = False
tTotal2.Text = 0
tSPPT2.Text = 0

End Sub

Private Sub cmdExit_Click()
On Error Resume Next
Unload Me
c_Ganti = 0
xID = ""
End Sub


Private Sub cmdNOP1_Click()
On Error GoTo Salah
J_Karakter
If Len(Trim(tNOP(0).Text)) - (totChar * 1) = 24 Then
    call_data
Else
    xID = 6
    frmLIST_Objek1.Show
End If
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cmdOK_Click()
On Error GoTo Salah
Dim Pesan
Screen.MousePointer = vbHourglass
If ccKec.Text = "" Then ccKec.Text = "*.*"
If ccKel.Text = "" Then ccKel.Text = "*.*"
    C_KEC = Left(Trim(ccKec.Text), 3)
    C_KEL = Left(Trim(ccKel.Text), 3)
    C_TAHUN = ccTahun.Text
    c_NOP = aNOP.Text
    
If J_CETAK = 1 Or J_CETAK = 100 Or J_CETAK = 400 Then
        If hTunggal.Value = 1 Then
            Pesan = "Apa anda yakin cetak SPPT secara tunggal?"
            C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
            " FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND ([SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP])='" & c_NOP & "' AND PROSES<>'N'" & _
            " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
        Else
            Pesan = "Apa anda yakin cetak SPPT secara massal?"
            If C_KEC = "*.*" And C_KEL = "*.*" Then
                If cRekam.Visible = True And cRekam.Value = 1 Then
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND (QOBJEKPAJAK.TGL_PEREKAMAN_OP > ='" & Format(dRekam1.Value, "YYYY/MM/DD") & "' AND QOBJEKPAJAK.TGL_PEREKAMAN_OP <='" & Format(dRekam2.Value, "YYYY/MM/DD") & "')" & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                Else
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                End If
            ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
                If cRekam.Visible = True And cRekam.Value = 1 Then
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND (QOBJEKPAJAK.TGL_PEREKAMAN_OP > ='" & Format(dRekam1.Value, "YYYY/MM/DD") & "' AND QOBJEKPAJAK.TGL_PEREKAMAN_OP <='" & Format(dRekam2.Value, "YYYY/MM/DD") & "')" & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                Else
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                End If
            Else
                If cRekam.Visible = True And cRekam.Value = 1 Then
                    C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                    ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND (QOBJEKPAJAK.TGL_PEREKAMAN_OP > ='" & Format(dRekam1.Value, "YYYY/MM/DD") & "' AND QOBJEKPAJAK.TGL_PEREKAMAN_OP <='" & Format(dRekam2.Value, "YYYY/MM/DD") & "')" & _
                    " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                    
                Else
                    C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                    ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                    " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                End If
            End If
        End If
        openDB (C_STR)
        
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    
    If rPajak.EOF Then
        MsgBox "SPPT Belum ditetapkan...", vbCritical, "Error1"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    i = 0
    Do While Not rPajak.EOF
    If rPajak!PROSES = "N" Then
        MsgBox "SPPT BELUM DITETAPKAN..!", vbCritical, "Tetnong"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    i = i + 1
    jum = jum + rPajak!PBB_YG_HARUS_DIBAYAR_SPPT
    T_Terbit = rPajak!TGL_TERBIT_SPPT
    rPajak.MoveNext
    Loop
    tSPPT.Text = Format(i, "#,#0")
    tTotal.Text = Format(jum, "#,#0")
    If hTunggal.Value = 0 Then
        dTerbit.Value = Format(T_Terbit, "DD/MM/YYYY")
    End If
    TANYA = MsgBox(Pesan, vbQuestion + vbYesNo, "Cetak")
    If TANYA = vbNo Then
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        sv_cetak
    End If
    rptPBB.Show
ElseIf J_CETAK = 2 Or J_CETAK = 200 Or J_CETAK = 500 Then
        If hTunggal.Value = 1 Then
            C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
            " FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND ([SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP])='" & c_NOP & "' AND PROSES<>'N'" & _
            " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
        Else
            If C_KEC = "*.*" And C_KEL = "*.*" Then
                If cRekam.Visible = True And cRekam.Value = 1 Then
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                " ,QOBJEKPAJAK.TGL_PEREKAMAN_OP FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND (QOBJEKPAJAK.TGL_PEREKAMAN_OP > ='" & Format(dRekam1.Value, "YYYY/MM/DD") & "' AND QOBJEKPAJAK.TGL_PEREKAMAN_OP <='" & Format(dRekam2.Value, "YYYY/MM/DD") & "')" & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                Else
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                " ,QOBJEKPAJAK.TGL_PEREKAMAN_OP FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                End If
            ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
                If cRekam.Visible = True And cRekam.Value = 1 Then
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                " ,QOBJEKPAJAK.TGL_PEREKAMAN_OP FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND (QOBJEKPAJAK.TGL_PEREKAMAN_OP > ='" & Format(dRekam1.Value, "YYYY/MM/DD") & "' AND QOBJEKPAJAK.TGL_PEREKAMAN_OP <='" & Format(dRekam2.Value, "YYYY/MM/DD") & "')" & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                Else
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                " ,QOBJEKPAJAK.TGL_PEREKAMAN_OP FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                End If
            Else
                If cRekam.Visible = True And cRekam.Value = 1 Then
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                ",QOBJEKPAJAK.TGL_PEREKAMAN_OP FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND (QOBJEKPAJAK.TGL_PEREKAMAN_OP > ='" & Format(dRekam1.Value, "YYYY/MM/DD") & "' AND QOBJEKPAJAK.TGL_PEREKAMAN_OP <='" & Format(dRekam2.Value, "YYYY/MM/DD") & "')" & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                Else
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                " ,QOBJEKPAJAK.TGL_PEREKAMAN_OP FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                End If
            End If
        End If
    openDB (C_STR)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        MsgBox "SPPT Belum ditetapkan...", vbCritical, "Error2"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    i = 0
    Do While Not rPajak.EOF
    If rPajak!PROSES = "N" Then
        MsgBox "SPPT BELUM DITETAPKAN..!", vbCritical, "Tetnong"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    i = i + 1
    'MsgBox rPajak!TGL_PEREKAMAN_OP
    jum = jum + rPajak!PBB_YG_HARUS_DIBAYAR_SPPT
    rPajak.MoveNext
    Loop
    tSPPT.Text = Format(i, "#,#0")
    tTotal.Text = Format(jum, "#,#0")
    rptPBB.Show
ElseIf J_CETAK = 3 Or J_CETAK = 300 Then
        SPPT_DHKP
ElseIf J_CETAK = 4 Then
    KLAS1
    rptPBB.Show
ElseIf J_CETAK = 5 Then
    tampil_Simulasi
    rptPBB.Show
ElseIf J_CETAK = 10 Or J_CETAK = 20 Or J_CETAK = 30 Or J_CETAK = 40 Or J_CETAK = 50 Or J_CETAK = 60 Then
    rptPBB.Show
ElseIf J_CETAK = 6 Then
    If K_JPB = "" Or K_JPB = 0 Then
        MsgBox "Pilih Jenis Penggunaan Bangunan...!", vbCritical, "Tetnong...!"
        
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'K_JPB = Left(Trim(cJPB.Text), 2)
    rptPBB.Show
End If
'
'
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub dJTempo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub cRekam_Click()
On Error Resume Next
If cRekam.Value = 1 Then
    cUrut.Value = 0
    tSPPT.Visible = False
    tSPPT2.Visible = False
    Label6.Visible = True
    Label10.Visible = True
    dRekam1.Visible = True
    dRekam2.Visible = True
    Label6.Visible = True
    Label10.Visible = True
    Label10.Caption = "Tanggal Rekam"
    'ccKel.Text = ""
Else
    'cUrut.Value = 1
    dRekam1.Visible = False
    dRekam2.Visible = False
    Label6.Visible = False
    Label10.Visible = False
    If J_CETAK = 1 Or J_CETAK = 2 Or J_CETAK = 100 Or J_CETAK = 200 Or J_CETAK = 400 Or J_CETAK = 500 Then
        tSPPT.Visible = True
        Label10.Visible = True
        Label10.Caption = "Jumlah SPPT"
    End If
        
    
    
End If

End Sub

Private Sub cRekam_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub cUrut_Click()
On Error Resume Next
If cUrut.Value = 1 Then
    cRekam.Value = 0
    tSPPT.Visible = True
    tSPPT2.Visible = True
    Label6.Visible = True
    Label10.Visible = True
    Label10.Caption = "[KDBlok].[NoUrut]"
    ccKel.Text = ""
    dRekam1.Visible = False
    dRekam2.Visible = False
Else
    'cRekam.Value = 1
    tSPPT.Text = 0
    tSPPT2.Text = 0
    tSPPT.Visible = False
    tSPPT2.Visible = False
    Label6.Visible = False
    Label10.Visible = False
    
End If
End Sub

Private Sub cUrut_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub dCetak_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub dRekam1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub dRekam2_KeyDown(KeyCode As Integer, Shift As Integer)
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
On Error GoTo Salah
Frame1.Left = 0

t_Normal

If c_Ganti = "" Or c_Ganti = 0 Then
ccTahun.Text = Format(Now, "yyyy")
ccTahun.Clear
ccTahun.Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    ccTahun.AddItem i
Next
dTerbit.Value = Format(Now, "dd/mm/yyyy")
dCetak.Value = Format(Now, "dd/mm/yyyy")
If J_CETAK = 60 Then
    ccKec.SetFocus
Else
    ccTahun.SetFocus
End If
If xID = "" Then
    hTunggal.Value = 0
    Frame4.Visible = False
    
'Else
 '   hTunggal.Value = 1
  '  Frame4.Visible = True
End If
If ccKec.ListCount <= 0 Then
    CALL_KEC
End If
'Frame4.Visible = False
'hTunggal.Value = 0

End If

If J_CETAK = 1 Or J_CETAK = 100 Or J_CETAK = 400 Then
    If J_CETAK = 400 Then
        Me.Caption = "Cetak PBB Secara Halaman Tunggal" & Format(Now, "yyyy")
    Else
        Me.Caption = "Cetak PBB Secara Massal Tahun " & Format(Now, "yyyy")
    End If
    
    tSPPT2.Visible = False: tSPPT.Visible = True: Label10.Visible = True: Label10.Caption = "Jumlah SPPT"
    dRekam1.Visible = False: dRekam2.Visible = False
    cRekam.Visible = True
            If cRekam.Value = 1 Then
                tSPPT2.Visible = False
                tSPPT.Visible = False
                Label6.Visible = True
                Label10.Visible = True
                Label10.Caption = "Tanggal Rekam"
                dRekam1.Visible = True
                dRekam2.Visible = True
            Else
                tSPPT.Enabled = True
                tSPPT2.Enabled = True
                Label6.Visible = False
                dRekam1.Visible = False
                dRekam2.Visible = False
            End If
    tTotal2.Visible = False
    If hTunggal.Value = 0 Then
        dTerbit.Enabled = False
    Else
        dTerbit.Enabled = True
    End If
ElseIf J_CETAK = 2 Or J_CETAK = 200 Or J_CETAK = 500 Then
    If J_CETAK = 500 Then
        Me.Caption = "Cetak SSPD Secara Halaman Tunggal Tahun " & Format(Now, "yyyy")
    Else
        Me.Caption = "Cetak SSPD Secara Massal Tahun " & Format(Now, "yyyy")
    End If
        tSPPT2.Visible = False: tSPPT.Visible = True: Label10.Visible = True: Label10.Caption = "Jumlah SPPT"
    dRekam1.Visible = False: dRekam2.Visible = False
    cRekam.Visible = True
            If cRekam.Value = 1 Then
                tSPPT2.Visible = False
                tSPPT.Visible = False
                Label6.Visible = True
                Label10.Visible = True
                Label10.Caption = "Tanggal Rekam"
                dRekam1.Visible = True
                dRekam2.Visible = True
            Else
                tSPPT.Enabled = True
                tSPPT2.Enabled = True
                Label6.Visible = False
                dRekam1.Visible = False
                dRekam2.Visible = False
            End If
    tTotal2.Visible = False
ElseIf J_CETAK = 3 Or J_CETAK = 4 Or J_CETAK = 10 Or J_CETAK = 20 Or J_CETAK = 30 Or J_CETAK = 40 Or J_CETAK = 50 Or J_CETAK = 60 Or J_CETAK = 300 Then
    ctk_Nilai
     tSPPT2.Visible = False
     dRekam1.Visible = False: dRekam2.Visible = False
    tTotal2.Visible = False
     dTerbit.Enabled = False
    dCetak.Enabled = False
    tNIP.Enabled = False
    Frame4.Visible = False
    hTunggal.Value = 0
    hTunggal.Visible = False
    
    If J_CETAK = 3 Or J_CETAK = 300 Then
        Label10.Caption = "[KDBlok].[NoUrut]"
        If xID = "" Then
            tSPPT2.Visible = False
            tSPPT.Visible = False
            tSPPT.Enabled = True
            tSPPT2.Enabled = True
            Label6.Visible = False
            Label10.Visible = False
        Else
            If cUrut.Value = 1 Then
                tSPPT2.Visible = True
                tSPPT.Visible = True
                tSPPT.Enabled = True
                tSPPT2.Enabled = True
                Label6.Visible = True
                Label10.Visible = True
                Label10.Caption = "[KDBlok].[NoUrut]"
                dRekam1.Visible = False
                dRekam2.Visible = False
            ElseIf cRekam.Value = 1 Then
                tSPPT2.Visible = False
                tSPPT.Visible = False
                Label6.Visible = True
                Label10.Visible = True
                Label10.Caption = "Tanggal Rekam"
                dRekam1.Visible = True
                dRekam2.Visible = True
            Else
                tSPPT2.Visible = False
                tSPPT.Visible = False
                tSPPT.Enabled = True
                tSPPT2.Enabled = True
                Label6.Visible = False
                Label10.Visible = False
                dRekam1.Visible = False
                dRekam2.Visible = False
            End If
        End If
        cUrut.Visible = True
        cRekam.Visible = True
        'cUrut.Value = 0
        tSPPT.Locked = False: tSPPT2.Locked = False
        Me.Caption = "Cetak Daftar Himpunan Ketetapan Pajak Tahun " & Format(Now, "yyyy")
        ccTahun.SetFocus
    ElseIf J_CETAK = 4 Then
        Me.Caption = "Cetak Klasifikasi dan NJOP Bumi " & Format(Now, "yyyy")
        ccTahun.SetFocus
    ElseIf J_CETAK = 10 Then
        Me.Caption = "Laporan Penilaian Individu " & Format(Now, "yyyy")
        ccTahun.SetFocus
    ElseIf J_CETAK = 20 Then
        Me.Caption = "Laporan Penilaian Bumi Secara Detail " & Format(Now, "yyyy")
        ccTahun.SetFocus
    ElseIf J_CETAK = 30 Then
        Me.Caption = "Laporan Penilaian Bangunan Secara Detail " & Format(Now, "yyyy")
        ccTahun.SetFocus
    ElseIf J_CETAK = 40 Then
        Me.Caption = "Laporan Penilaian Bumi dan Bangunan Tahun " & Format(Now, "yyyy")
        ccTahun.SetFocus
    ElseIf J_CETAK = 50 Then
        Me.Caption = "Laporan Penilaian Bumi dan Bangunan Tahun " & Format(Now, "yyyy") - 1
        ccTahun.Enabled = False
        ccKec.SetFocus
        ccKec.Text = "*.*"
    ElseIf J_CETAK = 60 Then
        Me.Caption = "Laporan Perbandingan Penilaian Bumi dan Bangunan"
        ccTahun.Enabled = False
        ccKec.SetFocus
        ccKec.Text = "*.*"
    End If
   
    
ElseIf J_CETAK = 5 Then
    ctk_Nilai
    Me.Caption = "Cetak Simulasi Perbandingan PBB Tahun " & Format(Now, "yyyy") & " dengan Tahun " & Val(Format(Now, "yyyy")) - 1
    dTerbit.Enabled = False
    dCetak.Enabled = False
    tNIP.Enabled = False
    tSPPT2.Visible = False 'True
    dRekam1.Visible = False: dRekam2.Visible = False
    tTotal2.Visible = False ' True
    Frame4.Visible = False
    hTunggal.Value = 0
    hTunggal.Visible = False
ElseIf J_CETAK = 6 Then
    'cJPB.Clear
    ctk_Nilai
    Me.Caption = "Cetak Daftar Biaya Komponen Bangunan Tahun " & Format(Now, "yyyy")
    
    tSPPT2.Visible = False
    dRekam1.Visible = False: dRekam2.Visible = False
    tTotal2.Visible = False
     dTerbit.Enabled = False
    dCetak.Enabled = False
    tNIP.Enabled = False
    Frame4.Visible = False
    hTunggal.Value = 0
    hTunggal.Visible = False
    ccTahun.SetFocus
    ccKec.Visible = False
    ccKel.Visible = False
    Label8.Visible = False
    
End If
If hTunggal.Value = 0 Then Frame4.Visible = False
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
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
On Error GoTo Salah
If ccKec.Text = "*.*" Then
    ccKel.Enabled = False
    ccKel.Text = "*.*"

Else
    ccKel.Enabled = True
    CALL_KEL
End If

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
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
            ccKel_Click
            Exit Sub
        End If
          If i = ccKel.ListCount - 1 Then
            If UCase(ccKel.List(i)) Like "*" + UCase(ccKel.Text) + "*" = False Then
                ccKel.Text = ccKel.List(0)
                ccKel_Click
                Exit Sub
            End If
        End If
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
c_Ganti = 0
xID = ""
End Sub

Private Sub hTunggal_Click()
On Error Resume Next
If hTunggal.Value = 1 Then
    Frame4.Visible = True
    ccKec.Enabled = False
    ccKel.Enabled = False
    'ccKec.Text = ""
    dTerbit.Enabled = True
Else
    Frame4.Visible = False
    ccKec.Enabled = True
    ccKel.Enabled = True
    dTerbit.Enabled = False
End If
End Sub

Private Sub hTunggal_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
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
If tNIP.Text = "" Or tNIP.Text = "-" Or tNIP.Text = "." Then
    tNIP.Text = 0
End If
tNIP.Alignment = 1

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

Private Sub tSPPT2_GotFocus()
On Error Resume Next
tSPPT2.SelStart = 0
tSPPT2.SelLength = Len(tSPPT2.Text)
tSPPT2.SetFocus
tSPPT2.Alignment = 0
End Sub

Private Sub tSPPT2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tSPPT2_LostFocus()
On Error Resume Next
If tSPPT2.Text = "" Or tSPPT2.Text = "-" Or tSPPT2.Text = "." Then
    tSPPT2.Text = 0
End If
tSPPT2.Alignment = 1
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

Sub sv_cetak()
On Error GoTo Salah
If hTunggal.Value = 1 Then
    U_STR = "UPDATE SPPT SET  STATUS_CETAK_SPPT='1',TGL_TERBIT_SPPT='" & Format(dTerbit.Value, "yyyy-mm-dd") & "',TGL_CETAK_SPPT='" & Format(dCetak.Value, "yyyy-mm-dd") & "',NIP_PENCETAK_SPPT='" & tNIP.Text & "' WHERE THN_PAJAK_SPPT='" & C_TAHUN & "' and ([KD_PROPINSI]+'.'+[KD_DATI2]+'.'+[KD_KECAMATAN]+'.'+[KD_KELURAHAN]+'.'+[KD_BLOK]+'-'+[NO_URUT]+'.'+[KD_JNS_OP])='" & c_NOP & "' "
Else
    If C_KEC = "*.*" And C_KEL = "*.*" Then
        U_STR = "UPDATE SPPT SET  STATUS_CETAK_SPPT='1',TGL_TERBIT_SPPT='" & Format(dTerbit.Value, "yyyy-mm-dd") & "',TGL_CETAK_SPPT='" & Format(dCetak.Value, "yyyy-mm-dd") & "',NIP_PENCETAK_SPPT='" & tNIP.Text & "' WHERE THN_PAJAK_SPPT='" & C_TAHUN & "'"
    ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
        U_STR = "UPDATE SPPT SET  STATUS_CETAK_SPPT='1',TGL_TERBIT_SPPT='" & Format(dTerbit.Value, "yyyy-mm-dd") & "',TGL_CETAK_SPPT='" & Format(dCetak.Value, "yyyy-mm-dd") & "',NIP_PENCETAK_SPPT='" & tNIP.Text & "' WHERE KD_KECAMATAN='" & C_KEC & "' AND THN_PAJAK_SPPT='" & C_TAHUN & "'"
    Else
        U_STR = "UPDATE SPPT SET  STATUS_CETAK_SPPT='1',TGL_TERBIT_SPPT='" & Format(dTerbit.Value, "yyyy-mm-dd") & "',TGL_CETAK_SPPT='" & Format(dCetak.Value, "yyyy-mm-dd") & "',NIP_PENCETAK_SPPT='" & tNIP.Text & "' WHERE KD_KECAMATAN='" & C_KEC & "' AND KD_KELURAHAN='" & C_KEL & "' AND THN_PAJAK_SPPT='" & C_TAHUN & "'"
    End If
End If
    openDB (U_STR)

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
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

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub call_data()
'If hTunggal.Value = 1 Then
On Error GoTo Salah
    C_STR = "Select * From SPPT where ([KD_PROPINSI]+'.'+[KD_DATI2]+'.'+[KD_KECAMATAN]+'.'+[KD_KELURAHAN]+'.'+[KD_BLOK]+'-'+[NO_URUT]+'.'+[KD_JNS_OP])='" & Trim(aNOP.Text) & "' AND THN_PAJAK_SPPT='" & ccTahun.Text & "' AND PROSES<>'N'"
'Else
'    C_sTR = "Select * From SPPT where KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(ccKel.Text), 3) & "' AND THN_PAJAK_SPPT='" & ccTahun.Text & "'AND PROSES<>'N'"
'End If
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If rPajak.EOF Then
    MsgBox "SPPT BELUM DITETAPKAN...", vbCritical, "TETNONG.."
    Exit Sub
End If

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub tampil_Simulasi()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
vOP.ListItems.Clear
'SPPT TAHUN BERJALAN
If hTunggal = 1 Then
        'c_str = "SELECT SPPT.PROSES,SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, SPPT.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP, SPPT.THN_PAJAK_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT FROM (SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_DATI2 = REF_KELURAHAN.KD_DATI2) AND (SPPT.KD_PROPINSI = REF_KELURAHAN.KD_PROPINSI) AND (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) where ([SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] ='" & aNOP.Text & "') AND (SPPT.THN_PAJAK_SPPT='" & C_TAHUN & "' OR SPPT.THN_PAJAK_SPPT='" & (C_TAHUN * 1) - 1 & "') ORDER BY SPPT.THN_PAJAK_SPPT ASC"
        'C_STR = "SELECT SPPT.KD_PROPINSI,SPPT.KD_DATI2,SPPT.KD_BLOK,SPPT.NO_URUT,SPPT.KD_JNS_OP,SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, SPPT.KD_KELURAHAN AS JUM_SPPT, REF_KELURAHAN.NM_KELURAHAN, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 & "')  and [SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] ='" & aNOP.Text & "' ) "
        C_STR = "SELECT SPPT.*, REF_KELURAHAN.NM_KELURAHAN,REF_KECAMATAN.NM_KECAMATAN FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) WHERE ([SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] ='" & aNOP.Text & "' ) AND SPPT.THN_PAJAK_SPPT='" & C_TAHUN * 1 & "'  ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
    Else
        If C_KEC = "*.*" And C_KEL = "*.*" Then
            'c_str = "SELECT SPPT.PROSES,SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, SPPT.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP, SPPT.THN_PAJAK_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT FROM (SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_DATI2 = REF_KELURAHAN.KD_DATI2) AND (SPPT.KD_PROPINSI = REF_KELURAHAN.KD_PROPINSI) AND (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) where SPPT.THN_PAJAK_SPPT='" & C_TAHUN & "' OR SPPT.THN_PAJAK_SPPT='" & (C_TAHUN * 1) - 1 & "' ORDER BY SPPT.THN_PAJAK_SPPT ASC"
            C_STR = "SELECT SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, Count(SPPT.KD_KELURAHAN) AS JUM_SPPT, REF_KELURAHAN.NM_KELURAHAN, Sum(SPPT.PBB_YG_HARUS_DIBAYAR_SPPT) AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT HAVING (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 & "'))  ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
            'c_STR = "SELECT KD_KECAMATAN,KD_KELURAHAN,THN_PAJAK_SPPT,PROSES,SUM(PBB_YG_HARUS_DIBAYAR_SPPT) AS JUM_PBB FROM SPPT where KD_KECAMATAN='" & C_KEC & "'  AND (THN_PAJAK_SPPT='" & C_TAHUN & "' GROUP BY KD_KECAMATAN,KD_KELURAHAN,THN_PAJAK_SPPT,PROSES,PBB_YG_HARUS_DIBAYAR_SPPT"
            C_STR = "SELECT SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, Count(SPPT.KD_KELURAHAN) AS JUM_SPPT, REF_KELURAHAN.NM_KELURAHAN, Sum(SPPT.PBB_YG_HARUS_DIBAYAR_SPPT) AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT HAVING (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 & "')  and SPPT.KD_KECAMATAN='" & C_KEC & "') ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"

            'c_STR = "SELECT SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, Sum(SPPT.PBB_YG_HARUS_DIBAYAR_SPPT) AS JLH_PBB, SPPT.THN_PAJAK_SPPT From SPPT GROUP BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT HAVING ((THN_PAJAK_SPPT='" & C_TAHUN & "' OR THN_PAJAK_SPPT='" & (C_TAHUN * 1) - 1 & "')) AND KD_KECAMATAN='" & C_KEC & "'"
        Else
            'c_str = "SELECT SPPT.PROSES,SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, SPPT.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP, SPPT.THN_PAJAK_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT FROM (SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_DATI2 = REF_KELURAHAN.KD_DATI2) AND (SPPT.KD_PROPINSI = REF_KELURAHAN.KD_PROPINSI) AND (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) where SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' AND (SPPT.THN_PAJAK_SPPT='" & C_TAHUN & "' OR SPPT.THN_PAJAK_SPPT='" & (C_TAHUN * 1) - 1 & "') ORDER BY SPPT.THN_PAJAK_SPPT ASC"
            'C_STR = "SELECT SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, COUNT(SPPT.KD_KELURAHAN) AS JUM_SPPT, REF_KELURAHAN.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, Sum(SPPT.PBB_YG_HARUS_DIBAYAR_SPPT) AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT HAVING (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 & "')and SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' ) "
            C_STR = "SELECT SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, COUNT(SPPT.KD_KELURAHAN) AS JUM_SPPT, REF_KELURAHAN.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, Sum(SPPT.PBB_YG_HARUS_DIBAYAR_SPPT) AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT HAVING (((SPPT.KD_KECAMATAN)='" & C_KEC & "') AND ((REF_KELURAHAN.KD_KELURAHAN)='" & C_KEL & "') AND ((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "'))ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        End If
    End If
    openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'  If rPajak.EOF Then
'        MsgBox "SPPT Belum ditetapkan...", vbCritical, "Error"
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    JUM1 = 0: JS1 = 0
    Do While Not rPajak.EOF
    If rPajak!PROSES <> "N" Or IsNull(rPajak!PROSES) = True Then
        If hTunggal.Value = 1 Then
            c1 = 1
            c2 = rPajak!PBB_YG_HARUS_DIBAYAR_SPPT
        Else
            c1 = rPajak!JUM_SPPT
            c2 = rPajak!JUM_PBB
        End If
        JUM1 = JUM1 + c2 'PBB_YG_HARUS_DIBAYAR_SPPT
        JS1 = JS1 + c1
     'JUM1 = JUM1 + rPajak!JUM_PBB 'PBB_YG_HARUS_DIBAYAR_SPPT
     'JS1 = JS1 + rPajak!JUM_SPPT
        i = i + 1
        
        vOP.ListItems.Add i, "", Format(i, "#")
        vOP.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        vOP.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![NM_KECAMATAN])
        vOP.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![NM_KELURAHAN])
        vOP.ListItems.Item(i).ListSubItems.Add 4, "", c1 'rPajak!JUM_SPPT
        vOP.ListItems.Item(i).ListSubItems.Add 5, "", c2 'rPajak!JUM_PBB
        vOP.ListItems.Item(i).ListSubItems.Add 6, "", 0
        vOP.ListItems.Item(i).ListSubItems.Add 7, "", 0
        vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!THN_PAJAK_SPPT
        vOP.ListItems.Item(i).ListSubItems.Add 9, "", 0
        vOP.ListItems.Item(i).ListSubItems.Add 10, "", 0
    End If
    rPajak.MoveNext
    Loop
    tTotal.Text = Format(JUM1, "#,#0")
    tSPPT.Text = Format(JS1, "#,#0")
    'tTotal2.Text = Format(JUM2, "#,#0")
'==================SPPT TAHUN SEBELUMNYA======================
If hTunggal = 1 Then
        'C_STR = "SELECT SPPT.KD_PROPINSI,SPPT.KD_DATI2,SPPT.KD_BLOK,SPPT.NO_URUT,SPPT.KD_JNS_OP,SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, SPPT.KD_KELURAHAN AS JUM_SPPT, REF_KELURAHAN.NM_KELURAHAN, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 - 1 & "')  and [SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] ='" & aNOP.Text & "' ) "
        'C_STR = "SELECT SPPT.*, REF_KELURAHAN.NM_KELURAHAN, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT WHERE (SPPT.THN_PAJAK_SPPT='" & C_TAHUN * 1 & "' and ([SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] ='" & aNOP.Text & "' )) "
        C_STR = "SELECT SPPT.*, REF_KELURAHAN.NM_KELURAHAN,REF_KECAMATAN.NM_KECAMATAN  FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) WHERE ([SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] ='" & aNOP.Text & "' ) AND SPPT.THN_PAJAK_SPPT='" & C_TAHUN * 1 - 1 & "'  ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
    Else
        If C_KEC = "*.*" And C_KEL = "*.*" Then
            C_STR = "SELECT SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, Count(SPPT.KD_KELURAHAN) AS JUM_SPPT, REF_KELURAHAN.NM_KELURAHAN, Sum(SPPT.PBB_YG_HARUS_DIBAYAR_SPPT) AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT HAVING (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 - 1 & "') )  ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
            C_STR = "SELECT SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, Count(SPPT.KD_KELURAHAN) AS JUM_SPPT, REF_KELURAHAN.NM_KELURAHAN, Sum(SPPT.PBB_YG_HARUS_DIBAYAR_SPPT) AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT HAVING (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 - 1 & "') and SPPT.KD_KECAMATAN='" & C_KEC & "') ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        Else
            'C_STR = "SELECT SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, COUNT(SPPT.KD_KELURAHAN) AS JUM_SPPT, REF_KELURAHAN.KD_KELURAHAN , REF_KELURAHAN.NM_KELURAHAN, Sum(SPPT.PBB_YG_HARUS_DIBAYAR_SPPT) AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT HAVING (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 - 1 & "')and SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' ) "
            C_STR = "SELECT SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, COUNT(SPPT.KD_KELURAHAN) AS JUM_SPPT, REF_KELURAHAN.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, Sum(SPPT.PBB_YG_HARUS_DIBAYAR_SPPT) AS JUM_PBB, SPPT.PROSES, SPPT.THN_PAJAK_SPPT FROM SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, SPPT.PROSES, SPPT.THN_PAJAK_SPPT HAVING (((SPPT.KD_KECAMATAN)='" & C_KEC & "') AND ((REF_KELURAHAN.KD_KELURAHAN)='" & C_KEL & "') AND ((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 - 1 & "'))ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        End If
    End If
    openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'  If rPajak.EOF Then
'        MsgBox "SPPT Belum ditetapkan...", vbCritical, "Error"
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
     C = 0: i = vOP.ListItems.Count: JS2 = 0: JUM1 = 0: JUM2 = 0
    Do While Not rPajak.EOF
    If rPajak!PROSES <> "N" Or IsNull(rPajak!PROSES) = True Then
        If hTunggal.Value = 1 Then
            c1 = 1
            c2 = rPajak!PBB_YG_HARUS_DIBAYAR_SPPT
        Else
            c1 = rPajak!JUM_SPPT
            c2 = rPajak!JUM_PBB
        End If
        JUM2 = JUM2 + c2 'PBB_YG_HARUS_DIBAYAR_SPPT
        JS2 = JS2 + c1
    
    i = i + 1
        vOP.ListItems.Add i, "", Format(i, "#")
        vOP.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        vOP.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![NM_KECAMATAN])
        vOP.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![NM_KELURAHAN])
        vOP.ListItems.Item(i).ListSubItems.Add 4, "", 0
        vOP.ListItems.Item(i).ListSubItems.Add 5, "", 0
        vOP.ListItems.Item(i).ListSubItems.Add 6, "", c1 'rPajak!JUM_SPPT
        vOP.ListItems.Item(i).ListSubItems.Add 7, "", c2 'rPajak!JUM_PBB
        vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!THN_PAJAK_SPPT
        vOP.ListItems.Item(i).ListSubItems.Add 9, "", 0
        vOP.ListItems.Item(i).ListSubItems.Add 10, "", 0
    End If
    rPajak.MoveNext
    Loop
    'tSPPT.Text = Format(C, "#,#0")
    'tTotal.Text = Format(JUM1, "#,#0")
    For J = 1 To vOP.ListItems.Count
        vOP.ListItems(J).ListSubItems(9).Text = vOP.ListItems(J).ListSubItems(4).Text * 1 + vOP.ListItems(J).ListSubItems(6).Text * 1
        vOP.ListItems(J).ListSubItems(10).Text = vOP.ListItems(J).ListSubItems(5).Text * 1 + vOP.ListItems(J).ListSubItems(7).Text * 1
        'vOP.ListItems(j).ListSubItems(8).Text = VOP1.ListItems(j).ListSubItems(4).Text * 1 + VOP1.ListItems(j).ListSubItems(6).Text * 1
    Next
    tSPPT2.Text = Format(JS2, "#,#0")
    tTotal2.Text = Format(JUM2, "#,#0")
    'JUM_SPPT1
    'JUM_SPPT2
    
    GABUNG
    SIMPAN_PBB

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub

Sub SIMPAN_PBB()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
c_DEL = "Delete From SPPT_SIMULASI1"
openDB (c_DEL)
For i = 1 To VOP1.ListItems.Count
    xNO = VOP1.ListItems.Item(i).ListSubItems(1).Text
    xKec = VOP1.ListItems.Item(i).ListSubItems(2).Text
    xKel = VOP1.ListItems.Item(i).ListSubItems(3).Text
    xSPPT1 = VOP1.ListItems.Item(i).ListSubItems(4).Text
    xPBB1 = VOP1.ListItems.Item(i).ListSubItems(5).Text
    xSPPT2 = VOP1.ListItems.Item(i).ListSubItems(6).Text
    xPBB2 = VOP1.ListItems.Item(i).ListSubItems(7).Text
    xTahun = ccTahun.Text 'VOP1.ListItems.Item(i).ListSubItems(8).Text
    c_ins = "INSERT INTO SPPT_SIMULASI1(NoUrut,KD_PROPINSI,KD_DATI2,NM_KECAMATAN,NM_KELURAHAN,SPPT1,PBB1,SPPT2,PBB2,THN_PAJAK) VALUES('" & xNO & "','12','12','" & xKec & "','" & xKel & "','" & xSPPT1 & "','" & xPBB1 & "','" & xSPPT2 & "','" & xPBB2 & "','" & xTahun & "')"
    openDB (c_ins)
Next

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub

Sub GABUNG()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
VOP1.ListItems.Clear
        If C_KEC = "*.*" And C_KEL = "*.*" Then
            C_KEC = "SELECT REF_KECAMATAN.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN FROM REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
            C_KEC = "SELECT REF_KECAMATAN.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN FROM REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI)WHERE (((REF_KECAMATAN.KD_KECAMATAN)='" & C_KEC & "')) ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        Else
            C_KEC = "SELECT REF_KECAMATAN.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN FROM REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI)WHERE (((REF_KECAMATAN.KD_KECAMATAN)='" & C_KEC & "')) AND (((REF_KELURAHAN.KD_KELURAHAN)='" & C_KEL & "'))ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        End If
    
    openDB (C_KEC)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
    i = i + 1
        VOP1.ListItems.Add i, "", Format(i, "#")
        VOP1.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        VOP1.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![NM_KECAMATAN])
        VOP1.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![NM_KELURAHAN])
        VOP1.ListItems.Item(i).ListSubItems.Add 4, "", 0 ' C1 'rPajak!JUM_SPPT
        VOP1.ListItems.Item(i).ListSubItems.Add 5, "", 0 'C2 'rPajak!JUM_PBB
        VOP1.ListItems.Item(i).ListSubItems.Add 6, "", 0
        VOP1.ListItems.Item(i).ListSubItems.Add 7, "", 0
        VOP1.ListItems.Item(i).ListSubItems.Add 8, "", 0 'rPajak!THN_PAJAK_SPPT
    rPajak.MoveNext
    Loop
    For J = 1 To VOP1.ListItems.Count
        c4 = 0: c5 = 0
        For K = 1 To vOP.ListItems.Count
           
            If vOP.ListItems(K).ListSubItems(2).Text = VOP1.ListItems(J).ListSubItems(2).Text And vOP.ListItems(K).ListSubItems(3).Text = VOP1.ListItems(J).ListSubItems(3).Text Then
                c4 = c4 + vOP.ListItems(K).ListSubItems(4).Text * 1
                c5 = c5 + vOP.ListItems(K).ListSubItems(5).Text * 1
                VOP1.ListItems(J).ListSubItems(4).Text = c4 'vOP.ListItems(k).ListSubItems(4).Text
                VOP1.ListItems(J).ListSubItems(5).Text = c5 'vOP.ListItems(k).ListSubItems(5).Text
                VOP1.ListItems(J).ListSubItems(6).Text = vOP.ListItems(K).ListSubItems(9).Text
                VOP1.ListItems(J).ListSubItems(7).Text = vOP.ListItems(K).ListSubItems(10).Text
                
            End If
        Next
    Next

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
    Screen.MousePointer = vbDefault
End Sub
Sub BAYAR_DHKP()
On Error GoTo Salah
If hTunggal = 1 Then
        C_STR = "SELECT PEMBAYARAN_SPPT.*, REF_KELURAHAN.NM_KELURAHAN,REF_KECAMATAN.NM_KECAMATAN FROM PEMBAYARAN_SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) WHERE ([PEMBAYARAN_SPPT].[KD_PROPINSI]+'.'+[PEMBAYARAN_SPPT].[KD_DATI2]+'.'+[PEMBAYARAN_SPPT].[KD_KECAMATAN]+'.'+[PEMBAYARAN_SPPT].[KD_KELURAHAN]+'.'+[PEMBAYARAN_SPPT].[KD_BLOK]+'-'+[PEMBAYARAN_SPPT].[NO_URUT]+'.'+[PEMBAYARAN_SPPT].[KD_JNS_OP] ='" & aNOP.Text & "' ) AND PEMBAYARAN_SPPT.THN_PAJAK_SPPT='" & C_TAHUN * 1 & "'  ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
Else
        If C_KEC = "*.*" And C_KEL = "*.*" Then
            C_STR = "SELECT PEMBAYARAN_SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, PEMBAYARAN_SPPT.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, PEMBAYARAN_SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, PEMBAYARAN_SPPT.PROSES, PEMBAYARAN_SPPT.THN_PAJAK_SPPT FROM PEMBAYARAN_SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (PEMBAYARAN_SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (PEMBAYARAN_SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, PEMBAYARAN_SPPT.THN_PAJAK_SPPT HAVING (((PEMBAYARAN_SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 & "'))  ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
            C_STR = "SELECT PEMBAYARAN_SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, PEMBAYARAN_SPPT.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, PEMBAYARAN_SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, PEMBAYARAN_SPPT.PROSES, PEMBAYARAN_SPPT.THN_PAJAK_SPPT FROM PEMBAYARAN_SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (PEMBAYARAN_SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (PEMBAYARAN_SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, PEMBAYARAN_SPPT.THN_PAJAK_SPPT HAVING (((PEMBAYARAN_SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 & "')  and PEMBAYARAN_SPPT.KD_KECAMATAN='" & C_KEC & "') ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        Else
            C_STR = "SELECT PEMBAYARAN_SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, PEMBAYARAN_SPPT.KD_KELURAHAN, REF_KELURAHAN.NM_KELURAHAN, PEMBAYARAN_SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, PEMBAYARAN_SPPT.PROSES, PEMBAYARAN_SPPT.THN_PAJAK_SPPT FROM PEMBAYARAN_SPPT INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN)) ON (PEMBAYARAN_SPPT.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (PEMBAYARAN_SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)GROUP BY SPPT.KD_KECAMATAN, REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN, PEMBAYARAN_SPPT.THN_PAJAK_SPPT HAVING (((PEMBAYARAN_SPPT.THN_PAJAK_SPPT)='" & C_TAHUN * 1 & "')  and (PEMBAYARAN_SPPT.KD_KECAMATAN='" & C_KEC & "' AND REF_KELURAHAN.KD_KELURAHAN='" & C_KEL & "') ORDER BY REF_KECAMATAN.NM_KECAMATAN,REF_KELURAHAN.NM_KELURAHAN ASC"
        End If
End If
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    i = i + 1
        vOP.ListItems.Add i, "", Format(i, "#")
        vOP.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        vOP.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![NM_KECAMATAN])
        vOP.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![NM_KELURAHAN])
        vOP.ListItems.Item(i).ListSubItems.Add 4, "", rPajak!Nm_wp
        vOP.ListItems.Item(i).ListSubItems.Add 5, "", c2 'rPajak!JUM_PBB
        vOP.ListItems.Item(i).ListSubItems.Add 6, "", 0
        vOP.ListItems.Item(i).ListSubItems.Add 7, "", 0
        vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!THN_PAJAK_SPPT
        vOP.ListItems.Item(i).ListSubItems.Add 9, "", 0
        vOP.ListItems.Item(i).ListSubItems.Add 10, "", 0
rPajak.MoveNext
Loop

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub SPPT_DHKP()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
vOP.ListItems.Clear
If hTunggal.Value = 1 Then
            C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
            " FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND ([SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP])='" & c_NOP & "' AND PROSES<>'N'" & _
            " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
        Else
            If C_KEC = "*.*" And C_KEL = "*.*" Then
                QQ = "1"
                If cRekam.Visible = True And cRekam.Value = 1 Then
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND (qobjekpajak.TGL_PEREKAMAN_OP > ='" & Format(dRekam1.Value, "YYYY/MM/DD") & "' AND qobjekpajak.TGL_PEREKAMAN_OP <='" & Format(dRekam2.Value, "YYYY/MM/DD") & "')" & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                Else
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                End If
            ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
                QQ = "2"
                If cRekam.Visible = True And cRekam.Value = 1 Then
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND (qobjekpajak.TGL_PEREKAMAN_OP > ='" & Format(dRekam1.Value, "YYYY/MM/DD") & "' AND qobjekpajak.TGL_PEREKAMAN_OP <='" & Format(dRekam2.Value, "YYYY/MM/DD") & "')" & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                Else
                C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                End If
            Else
                QQ = "3"
                If cUrut.Visible = True And cUrut.Value = 1 Then
                'C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                " FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' AND (SPPT.NO_URUT*1>='" & tSPPT.Text * 1 & "' AND SPPT.NO_URUT*1 <='" & tSPPT2.Text * 1 & "') AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                    If Len(Trim(tSPPT.Text)) = 4 Then
                        C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                        ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' AND (SPPT.NO_URUT>='" & tSPPT.Text & "' AND SPPT.NO_URUT <='" & tSPPT2.Text & "') AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                        " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                    Else
                        C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                        ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' AND ((SPPT.KD_BLOK> ='" & Left(tSPPT.Text, 3) & "' AND SPPT.KD_BLOK< ='" & Left(tSPPT2.Text, 3) & "' ) AND SPPT.NO_URUT>='" & Right(tSPPT.Text, 4) & "' AND SPPT.NO_URUT <='" & Right(tSPPT2.Text, 4) & "') AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                        " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                    End If
                ElseIf cRekam.Visible = True And cRekam.Value = 1 Then
                    C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                    ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) AND (qobjekpajak.TGL_PEREKAMAN_OP > ='" & Format(dRekam1.Value, "YYYY/MM/DD") & "' AND qobjekpajak.TGL_PEREKAMAN_OP <='" & Format(dRekam2.Value, "YYYY/MM/DD") & "')" & _
                    " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                    
                Else
                    C_STR = "SELECT REF_MAP.KD_MAP, REF_JNS_SEKTOR.NM_SEKTOR, SPPT.PROSES,[SPPT].[KD_PROPINSI]+'.'+[SPPT].[KD_DATI2]+'.'+[SPPT].[KD_KECAMATAN]+'.'+[SPPT].[KD_KELURAHAN]+'.'+[SPPT].[KD_BLOK]+'-'+[SPPT].[NO_URUT]+'.'+[SPPT].[KD_JNS_OP] AS NOPQ, SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.THN_PAJAK_SPPT, SPPT.NM_WP_SPPT, SPPT.JLN_WP_SPPT, SPPT.BLOK_KAV_NO_WP_SPPT, SPPT.RW_WP_SPPT, SPPT.RT_WP_SPPT, SPPT.KELURAHAN_WP_SPPT, SPPT.KOTA_WP_SPPT, SPPT.KD_POS_WP_SPPT, SPPT.NPWP_SPPT, SPPT.NO_PERSIL_SPPT, SPPT.KD_KLS_TANAH, SPPT.KD_KLS_BNG, SPPT.LUAS_BUMI_SPPT, SPPT.LUAS_BNG_SPPT, SPPT.NJOP_BUMI_SPPT, SPPT.NJOP_BNG_SPPT, SPPT.NJOP_SPPT, SPPT.NJOPTKP_SPPT, SPPT.NJKP_SPPT, SPPT.PBB_TERHUTANG_SPPT, SPPT.FAKTOR_PENGURANG_SPPT, SPPT.PBB_YG_HARUS_DIBAYAR_SPPT, SPPT.TGL_JATUH_TEMPO_SPPT, SPPT.TGL_TERBIT_SPPT, SPPT.TGL_CETAK_SPPT, TEMPAT_BAYAR.NM_TP, QOBJEKPAJAK.JALAN_OP, QOBJEKPAJAK.BLOK_KAV_NO_OP, QOBJEKPAJAK.RW_OP, QOBJEKPAJAK.RT_OP, QOBJEKPAJAK.NM_KECAMATAN, QOBJEKPAJAK.NM_KELURAHAN" & _
                    ",QOBJEKPAJAK.TGL_PEREKAMAN_OP  FROM QOBJEKPAJAK INNER JOIN (((SPPT INNER JOIN REF_KELURAHAN ON (SPPT.KD_KECAMATAN = REF_KELURAHAN.KD_KECAMATAN) AND (SPPT.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN)) INNER JOIN (REF_JNS_SEKTOR INNER JOIN REF_MAP ON REF_JNS_SEKTOR.KD_SEKTOR = REF_MAP.KD_SEKTOR) ON REF_KELURAHAN.KD_SEKTOR = REF_MAP.KD_SEKTOR) INNER JOIN TEMPAT_BAYAR ON SPPT.KD_TP = TEMPAT_BAYAR.KD_TP) ON (QOBJEKPAJAK.KD_JNS_OP = SPPT.KD_JNS_OP) AND (QOBJEKPAJAK.NO_URUT = SPPT.NO_URUT) AND (QOBJEKPAJAK.KD_BLOK = SPPT.KD_BLOK) AND (QOBJEKPAJAK.KD_KELURAHAN = SPPT.KD_KELURAHAN) AND (QOBJEKPAJAK.KD_Kecamatan = SPPT.KD_KECAMATAN) WHERE SPPT.KD_KECAMATAN='" & C_KEC & "' AND SPPT.KD_KELURAHAN='" & C_KEL & "' AND (((SPPT.THN_PAJAK_SPPT)='" & C_TAHUN & "')) " & _
                    " ORDER BY SPPT.KD_KECAMATAN, SPPT.KD_KELURAHAN, SPPT.KD_BLOK, SPPT.NO_URUT, SPPT.KD_JNS_OP"
                End If
            End If
        End If
    openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
xRec = rPajak.RecordCount
'    If rPajak.EOF Then
'        MsgBox "SPPT Belum ditetapkan...", vbCritical, "Error3"
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    i = 0
    Do While Not rPajak.EOF
    If rPajak!PROSES = "N" Then
        MsgBox "SPPT BELUM DITETAPKAN..!", vbCritical, "Tetnong"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
            i = i + 1
            Me.Caption = "Proses Objek Pajak : " & Round(i / xRec * 100, 0) & "%"
            vOP.ListItems.Add i, "", Format(i, "#")
            vOP.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
            vOP.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!KD_KECAMATAN
            vOP.ListItems.Item(i).ListSubItems.Add 3, "", rPajak!NM_KECAMATAN
            vOP.ListItems.Item(i).ListSubItems.Add 4, "", rPajak!KD_KELURAHAN
            vOP.ListItems.Item(i).ListSubItems.Add 5, "", rPajak!NM_KELURAHAN
            vOP.ListItems.Item(i).ListSubItems.Add 6, "", rPajak!NOPQ
            vOP.ListItems.Item(i).ListSubItems.Add 7, "", rPajak!NM_WP_SPPT
            If rPajak!JLN_WP_SPPT = "-" Or IsNull(rPajak!JLN_WP_SPPT) = True Then
                If rPajak!KOTA_WP_SPPT = "PAKPAK BHARAT" Then
                    vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!KELURAHAN_WP_SPPT & "(!)"
                Else
                    vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!KELURAHAN_WP_SPPT & "-" & rPajak!KOTA_WP_SPPT & "(!)"
                End If
            Else
                If UCase(Trim(rPajak!KOTA_WP_SPPT)) = "PAKPAK BHARAT" Then
                    'vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!JLN_WP_SPPT & ", " & rPajak!KELURAHAN_WP_SPPT
                    
                   ' If UCase(Trim(rPajak!KELURAHAN_WP_SPPT)) = UCase(Trim(rPajak!JLN_WP_SPPT)) Then
                   '     vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!JLN_WP_SPPT '& "-" & rPajak!KELURAHAN_WP_SPPT
                   ' Else
                        vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!JLN_WP_SPPT '& "-" & rPajak!KELURAHAN_WP_SPPT
                   ' End If
                Else
                    vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!JLN_WP_SPPT & "-" & rPajak!KOTA_WP_SPPT
                End If
            End If
            If Right(vOP.ListItems.Item(i).ListSubItems(8).Text, 1) = "-" Then vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!JLN_WP_SPPT
            vOP.ListItems.Item(i).ListSubItems.Add 9, "", rPajak!JALAN_OP
            vOP.ListItems.Item(i).ListSubItems.Add 10, "", rPajak!NM_SEKTOR
            vOP.ListItems.Item(i).ListSubItems.Add 11, "", rPajak!PBB_YG_HARUS_DIBAYAR_SPPT
            vOP.ListItems.Item(i).ListSubItems.Add 12, "", rPajak!TGL_TERBIT_SPPT
            vOP.ListItems.Item(i).ListSubItems.Add 13, "", C_TAHUN
            vOP.ListItems.Item(i).ListSubItems.Add 14, "", rPajak!NM_TP
            vOP.ListItems.Item(i).ListSubItems.Add 15, "", 0 'BUKU
            vOP.ListItems.Item(i).ListSubItems.Add 16, "", 0 'JUMLAH DIBAYARKAN
            vOP.ListItems.Item(i).ListSubItems.Add 17, "", 0 'TGL PEMBAYARAN
            vOP.ListItems.Item(i).ListSubItems.Add 18, "", rPajak!PBB_YG_HARUS_DIBAYAR_SPPT
            vOP.ListItems.Item(i).ListSubItems.Add 19, "", rPajak!LUAS_BUMI_SPPT
            vOP.ListItems.Item(i).ListSubItems.Add 20, "", rPajak!LUAS_BNG_SPPT
            If IsNull(rPajak!KOTA_WP_SPPT) = True Or Trim(rPajak!KOTA_WP_SPPT) = "-" Then
                vOP.ListItems.Item(i).ListSubItems.Add 21, "", "-"
            Else
                vOP.ListItems.Item(i).ListSubItems.Add 21, "", rPajak!KOTA_WP_SPPT
            End If
            jum = jum + rPajak!PBB_YG_HARUS_DIBAYAR_SPPT
    
rPajak.MoveNext
Loop
    'tSPPT.Text = Format(Q, "#,#0")
    tTotal.Text = Format(jum, "#,#0")
    vOP.ColumnHeaders(2).Text = "NO"
    vOP.ColumnHeaders(3).Text = "KODE"
    vOP.ColumnHeaders(4).Text = "KECAMATAN"
    vOP.ColumnHeaders(5).Text = "KODE"
    vOP.ColumnHeaders(6).Text = "KELURAHAN"
    vOP.ColumnHeaders(7).Text = "NOP"
    vOP.ColumnHeaders(8).Text = "NAMA WP"
    vOP.ColumnHeaders(9).Text = "ALAMAT WP"
    vOP.ColumnHeaders(10).Text = "ALAMAT OP"
    vOP.ColumnHeaders(11).Text = "SEKTOR"
    vOP.ColumnHeaders(12).Text = "PBB"
    vOP.ColumnHeaders(13).Text = "TGL TERBIT"
    vOP.ColumnHeaders(14).Text = "TAHUN"
    vOP.ColumnHeaders(15).Text = "TEMPAT"
    vOP.ColumnHeaders(16).Text = "BUKU"
    vOP.ColumnHeaders(17).Text = "BAYAR"
    vOP.ColumnHeaders(18).Text = "T_BAYAR"
    vOP.ColumnHeaders(19).Text = "JUM_AKHIR"
    BUKU
    BAYAR
    SIMPAN_DHKP
Me.Caption = "Cetak SPPT Secara Massal : Sukses!"
rptPBB.Show

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub

Sub BUKU()
On Error GoTo Salah
C_OBJ = "SELECT * FROM REF_BUKU "
openDB (C_OBJ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
xRec = vOP.ListItems.Count '* rPajak.RecordCount
K = 0
Do While Not rPajak.EOF
    K = K + 1
    

    'xBUKU(k) = rPajak!KD_BUKU
    'xMIN(k) = rPajak!NILAI_MIN_BUKU
    'xMax(k) = rPajak!NILAI_MAX_BUKU
    For i = 1 To vOP.ListItems.Count
        If vOP.ListItems.Item(i).ListSubItems(11).Text * 1 >= rPajak!NILAI_MIN_BUKU And vOP.ListItems.Item(i).ListSubItems(11).Text * 1 <= rPajak!NILAI_MAX_BUKU Then
            Me.Caption = "Proses Penentuan Jenis Buku : " & Round(i / xRec * 100, 0) & "%"
            vOP.ListItems.Item(i).ListSubItems(15).Text = "BUKU " & rPajak!KD_BUKU
        End If
    Next

rPajak.MoveNext
Loop
'For i = 1 To vOP.ListItems.Count
'        'If vOP.ListItems.Item(i).ListSubItems(11).Text * 1 >= rPajak!NILAI_MIN_BUKU And vOP.ListItems.Item(i).ListSubItems(11).Text * 1 <= rPajak!NILAI_MAX_BUKU Then
'        '    vOP.ListItems.Item(i).ListSubItems(15).Text rPajak!KD_BUKU
'        'End If
'        If vOP.ListItems.Item(i).ListSubItems(11).Text * 1 >= xMIN(1) And vOP.ListItems.Item(i).ListSubItems(11).Text * 1 <= xMax(1) Then
'            vOP.ListItems.Item(i).ListSubItems(15).Text = xBUKU(1)
'        ElseIf vOP.ListItems.Item(i).ListSubItems(11).Text * 1 >= xMIN(2) And vOP.ListItems.Item(i).ListSubItems(11).Text * 1 <= xMax(2) Then
'            vOP.ListItems.Item(i).ListSubItems(15).Text = xBUKU(2)
'        ElseIf vOP.ListItems.Item(i).ListSubItems(11).Text * 1 >= xMIN(3) And vOP.ListItems.Item(i).ListSubItems(11).Text * 1 <= xMax(3) Then
'            vOP.ListItems.Item(i).ListSubItems(15).Text = xBUKU(3)
'        ElseIf vOP.ListItems.Item(i).ListSubItems(11).Text * 1 >= xMIN(4) And vOP.ListItems.Item(i).ListSubItems(11).Text * 1 <= xMax(4) Then
'            vOP.ListItems.Item(i).ListSubItems(15).Text = xBUKU(4)
'        ElseIf vOP.ListItems.Item(i).ListSubItems(11).Text * 1 >= xMIN(5) And vOP.ListItems.Item(i).ListSubItems(11).Text * 1 <= xMax(5) Then
'            vOP.ListItems.Item(i).ListSubItems(15).Text = xBUKU(5)
'        End If
'Next

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub BAYAR()
On Error GoTo Salah
If hTunggal.Value = 1 Then
            C_BYR = "SELECT * FROM PEMBAYARAN_SPPT WHERE THN_PAJAK_SPPT='" & C_TAHUN & "' AND (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & c_NOP & "' "
Else
            If C_KEC = "*.*" And C_KEL = "*.*" Then
                C_BYR = "SELECT * FROM PEMBAYARAN_SPPT WHERE THN_PAJAK_SPPT='" & C_TAHUN & "' "
            ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
                C_BYR = "SELECT * FROM PEMBAYARAN_SPPT WHERE THN_PAJAK_SPPT='" & C_TAHUN & "' AND KD_KECAMATAN='" & C_KEC & "'"
                
            Else
                C_BYR = "SELECT * FROM PEMBAYARAN_SPPT WHERE THN_PAJAK_SPPT='" & C_TAHUN & "' AND KD_KECAMATAN='" & C_KEC & "' AND KD_KELURAHAN='" & C_KEL & "'"
            End If
End If
openDB (C_BYR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
xRec = rPajak.RecordCount
C = 0
Do While Not rPajak.EOF
C = C + 1
        Me.Caption = "Proses Objek Pajak yang sudah dibayar : " & Round(C / xRec * 100, 0) & "%"
    For i = 1 To vOP.ListItems.Count
        If rPajak!KD_PROPINSI & "." & rPajak!KD_DATI2 & "." & rPajak!KD_KECAMATAN & "." & rPajak!KD_KELURAHAN & "." & rPajak!KD_BLOK & "-" & rPajak!NO_URUT & "." & rPajak!KD_JNS_OP = Trim(vOP.ListItems.Item(i).ListSubItems(6).Text) Then
            vOP.ListItems.Item(i).ListSubItems(16).Text = rPajak!JML_SPPT_YG_DIBAYAR
            vOP.ListItems.Item(i).ListSubItems(17).Text = rPajak!TGL_PEMBAYARAN_SPPT
            vOP.ListItems.Item(i).ListSubItems(11).Text = vOP.ListItems.Item(i).ListSubItems(11).Text * 1 - (vOP.ListItems.Item(i).ListSubItems(16).Text * 1)
            If vOP.ListItems.Item(i).ListSubItems(11).Text <= 0 Then vOP.ListItems.Item(i).ListSubItems(11).Text = 0
        'Else
            'vOP.ListItems.Item(i).ListSubItems(18).Text = vOP.ListItems.Item(i).ListSubItems(11).Text * 1 - (vOP.ListItems.Item(i).ListSubItems(16).Text * 1)
        End If
    Next
rPajak.MoveNext
Loop

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub SIMPAN_DHKP()
'On Error GoTo Salah
On Error Resume Next
Screen.MousePointer = vbHourglass
c_DEL = "Delete  From TEMP_DHKP"
openDB (c_DEL)
'c_ins = "select * from TEMP_DHKP"
'openDB (c_ins)
'If rPajak.RecordCount > 0 Then rPajak.MoveFirst
xRec = vOP.ListItems.Count
For i = 1 To vOP.ListItems.Count
    Me.Caption = "Simpan Objek Pajak dan Persiapan Pencetakan: " & Round(i / xRec * 100, 0) & "%"
    xProp = "12": xKab = "12"
    cKode1 = vOP.ListItems.Item(i).ListSubItems(2).Text
    xKec = vOP.ListItems.Item(i).ListSubItems(3).Text
    cKode2 = vOP.ListItems.Item(i).ListSubItems(4).Text
    xKel = vOP.ListItems.Item(i).ListSubItems(5).Text
    xNOP = vOP.ListItems.Item(i).ListSubItems(6).Text
    xNama = vOP.ListItems.Item(i).ListSubItems(7).Text
    xAlamat = vOP.ListItems.Item(i).ListSubItems(8).Text
    cAlamat = vOP.ListItems.Item(i).ListSubItems(9).Text
    xSektor = vOP.ListItems.Item(i).ListSubItems(10).Text
    xPBB = vOP.ListItems.Item(i).ListSubItems(11).Text
    xTerbit = vOP.ListItems.Item(i).ListSubItems(12).Text
    xTahun = vOP.ListItems.Item(i).ListSubItems(13).Text
    xTempat = vOP.ListItems.Item(i).ListSubItems(14).Text
    xBUKU = vOP.ListItems.Item(i).ListSubItems(15).Text
    xBayar = Val(vOP.ListItems.Item(i).ListSubItems(16).Text)
    xTanggal = vOP.ListItems.Item(i).ListSubItems(17).Text
    If xTanggal = 0 Then xTanggal = ""
    cBayar = Val(vOP.ListItems.Item(i).ListSubItems(18).Text)
    tLuas = vOP.ListItems.Item(i).ListSubItems(19).Text
    bLuas = vOP.ListItems.Item(i).ListSubItems(20).Text
    cKota = vOP.ListItems.Item(i).ListSubItems(21).Text
        c_ins = "INSERT INTO TEMP_DHKP(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,NM_KECAMATAN,KD_KELURAHAN,NM_KELURAHAN,NOPQ,NM_WP_SPPT,JLN_WP_SPPT,JALAN_OP,NM_SEKTOR,PBB_YG_HARUS_DIBAYAR_SPPT,TGL_TERBIT_SPPT,THN_PAJAK,NM_TP,BUKU,JUM_BAYAR,TGL_BAYAR,JUM_AKHIR,LUAS_BUMI_SPPT,LUAS_BNG_SPPT,KOTA_WP)" & _
                " VALUES('" & xProp & "','" & xKab & "','" & cKode1 & "','" & xKec & "','" & cKode2 & "','" & xKel & "','" & xNOP & "','" & xNama & "','" & xAlamat & "','" & cAlamat & "','" & xSektor & "','" & xPBB & "','" & Format(xTerbit, "yyyy-mm-dd") & "','" & xTahun & "','" & xTempat & "','" & xBUKU & "','" & xBayar & "','" & Format(xTanggal, "yyyy-mm-dd") & "','" & cBayar & "','" & tLuas & "','" & bLuas & "','" & cKota & "')"
        openDB (c_ins)

'    rPajak.AddNew
'        rPajak!KD_PROPINSI = xProp
'        rPajak!KD_DATI2 = xKab
'        rPajak!KD_KECAMATAN = cKode1
'        rPajak!NM_KECAMATAN = xKec
'        rPajak!KD_KELURAHAN = cKode2
'        rPajak!NM_KELURAHAN = xKel
'        rPajak!NOPQ = xNOP
'        rPajak!NM_WP_SPPT = xNama
'        rPajak!JLN_WP_SPPT = xAlamat
'        rPajak!JALAN_OP = cAlamat
'        rPajak!NM_SEKTOR = xSektor
'        rPajak!PBB_YG_HARUS_DIBAYAR_SPPT = xPBB
'        rPajak!TGL_TERBIT_SPPT = xTerbit
'        rPajak!THN_PAJAK = xTahun
'        rPajak!NM_TP = xTempat
'        rPajak!BUKU = xBUKU
'        rPajak!JUM_BAYAR = xBayar
'        rPajak!TGL_BAYAR = xTanggal
'        rPajak!JUM_AKHIR = cBayar
'        rPajak!LUAS_BUMI_SPPT = tLuas
'        rPajak!LUAS_BNG_SPPT = bLuas
'    rPajak.Update
Next
i_DHKP = "SP_DHKP_SPPT '" & QQ & "','" & ccTahun.Text & "','" & Left(Trim(ccKec.Text), 3) & "','" & Left(Trim(ccKel.Text), 3) & "'"
openDB (i_DHKP)
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'Salah:
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub

Sub t_Normal()
On Error Resume Next
Me.Height = 4380
Me.Width = 6000
Picture1.Top = 3315
cmdOK.Top = 3480
cmdCear.Top = 3480
cmdExit.Top = 3480
Frame1.Visible = False
Label6.Visible = False
cUrut.Visible = False
cRekam.Visible = False
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
End Sub
Sub ctk_Nilai()
On Error Resume Next
tTotal.Visible = True
    Label4.Visible = True
    tSPPT2.Visible = False
    Label6.Visible = False
    dRekam1.Visible = False
    dRekam2.Visible = False
If J_CETAK = 6 Then
    Me.Height = 5580
    Me.Width = 6000
    Picture1.Top = 4440
    cmdOK.Top = 4590
    cmdCear.Top = 4590
    cmdExit.Top = 4590
    Frame1.Visible = True
    'Label10.Caption = "JPB"
ElseIf J_CETAK = 3 Or J_CETAK = 300 Then
    Frame1.Visible = False
    Me.Height = 3125
    Me.Width = 6000
    Picture1.Top = 2045
    cmdOK.Top = 2255
    cmdCear.Top = 2255
    cmdExit.Top = 2255
    Label10.Caption = "Nomor SPPT"
    tTotal.Visible = False
    Label4.Visible = False
    tSPPT2.Visible = True
    Label6.Visible = True
    dRekam1.Visible = True
    dRekam2.Visible = True
Else
    Frame1.Visible = False
    Me.Height = 2625
    Me.Width = 6000
    Picture1.Top = 1545
    cmdOK.Top = 1755
    cmdCear.Top = 1755
    cmdExit.Top = 1755
    Label10.Caption = "Jumlah SPPT"
End If
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2

End Sub
Sub KLAS1()
On Error GoTo Salah
vOP.ListItems.Clear
If C_KEC = "*.*" And C_KEL = "*.*" Then
            C_STR = "SELECT JALAN.*, DAT_NIR.THN_NIR_ZNT, DAT_NIR.[NIR], REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN FROM (JALAN INNER JOIN DAT_NIR ON (JALAN.KD_ZNT = DAT_NIR.KD_ZNT) AND (JALAN.KD_KELURAHAN = DAT_NIR.KD_KELURAHAN) AND (JALAN.KD_KECAMATAN = DAT_NIR.KD_KECAMATAN)) INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI)) ON (JALAN.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) AND (JALAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) WHERE DAT_NIR.THN_NIR_ZNT='" & C_TAHUN & "'"
        ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
            C_STR = "SELECT JALAN.*, DAT_NIR.THN_NIR_ZNT, DAT_NIR.[NIR], REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN FROM (JALAN INNER JOIN DAT_NIR ON (JALAN.KD_ZNT = DAT_NIR.KD_ZNT) AND (JALAN.KD_KELURAHAN = DAT_NIR.KD_KELURAHAN) AND (JALAN.KD_KECAMATAN = DAT_NIR.KD_KECAMATAN)) INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI)) ON (JALAN.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) AND (JALAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) WHERE DAT_NIR.THN_NIR_ZNT='" & C_TAHUN & "' AND JALAN.KD_KECAMATAN='" & C_KEC & "'"
        Else
            C_STR = "SELECT JALAN.*, DAT_NIR.THN_NIR_ZNT, DAT_NIR.[NIR], REF_KECAMATAN.NM_KECAMATAN, REF_KELURAHAN.NM_KELURAHAN FROM (JALAN INNER JOIN DAT_NIR ON (JALAN.KD_ZNT = DAT_NIR.KD_ZNT) AND (JALAN.KD_KELURAHAN = DAT_NIR.KD_KELURAHAN) AND (JALAN.KD_KECAMATAN = DAT_NIR.KD_KECAMATAN)) INNER JOIN (REF_KELURAHAN INNER JOIN REF_KECAMATAN ON (REF_KELURAHAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) AND (REF_KELURAHAN.KD_DATI2 = REF_KECAMATAN.KD_DATI2) AND (REF_KELURAHAN.KD_PROPINSI = REF_KECAMATAN.KD_PROPINSI)) ON (JALAN.KD_KELURAHAN = REF_KELURAHAN.KD_KELURAHAN) AND (JALAN.KD_KECAMATAN = REF_KECAMATAN.KD_KECAMATAN) WHERE DAT_NIR.THN_NIR_ZNT='" & C_TAHUN & "' AND JALAN.KD_KECAMATAN='" & C_KEC & "' AND JALAN.KD_KELURAHAN='" & C_KEL & "'"
        End If
openDB (C_STR)
xRec = rPajak.RecordCount
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    i = i + 1
        Me.Caption = "Proses Nama Jalan : " & Round(i / xRec * 100, 0) & "%"
        vOP.ListItems.Add i, "", Format(i, "#")
        vOP.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        vOP.ListItems.Item(i).ListSubItems.Add 2, "", rPajak!KD_KECAMATAN
        vOP.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![NM_KECAMATAN])
        vOP.ListItems.Item(i).ListSubItems.Add 4, "", rPajak![KD_KELURAHAN]
        vOP.ListItems.Item(i).ListSubItems.Add 5, "", Trim(rPajak![NM_KELURAHAN])
        vOP.ListItems.Item(i).ListSubItems.Add 6, "", rPajak!NM_JLN
        vOP.ListItems.Item(i).ListSubItems.Add 7, "", rPajak!NIR
        vOP.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!KD_ZNT
        vOP.ListItems.Item(i).ListSubItems.Add 9, "", ccTahun.Text
        vOP.ListItems.Item(i).ListSubItems.Add 10, "", 0 'Kelas Bumi
        vOP.ListItems.Item(i).ListSubItems.Add 11, "", 0 'Nilai Min
        vOP.ListItems.Item(i).ListSubItems.Add 12, "", 0 'Nilai Max
        vOP.ListItems.Item(i).ListSubItems.Add 13, "", 0 'NJOP Bumi
        vOP.ListItems.Item(i).ListSubItems.Add 14, "", rPajak!KD_BLOK
rPajak.MoveNext
Loop
    c_KLAS1
    SIMPAN_KLAS1
    Me.Caption = "Proses Persiapan Pencetakan Klasifikasi NJOP Tanah Sukses!"

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub c_KLAS1()
On Error GoTo Salah
Dim K_TAHUN
C_STR = "SELECT THN_AWAL_KLS_TANAH  FROM KELAS_TANAH ORDER BY THN_AWAL_KLS_TANAH DESC"
openDB (C_STR)
K_TAHUN = rPajak!THN_AWAL_KLS_TANAH
'MsgBox K_TAHUN
C_STR = "SELECT * FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH ='" & K_TAHUN & "' ORDER BY KD_KLS_TANAH ASC,THN_AWAL_KLS_TANAH DESC"
openDB (C_STR)
xRec = rPajak.RecordCount '* vOP.ListItems.Count
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
J = 0
Do While Not rPajak.EOF
    J = J + 1
    Me.Caption = "Proses Klasifikasi Tanah : " & Round(J / xRec * 100, 0) & "%"
    For i = 1 To vOP.ListItems.Count
        
        If vOP.ListItems.Item(i).ListSubItems(7).Text * 1 >= rPajak!NILAI_MIN_TANAH And vOP.ListItems.Item(i).ListSubItems(7).Text * 1 <= rPajak!NILAI_MAX_TANAH Then
            vOP.ListItems.Item(i).ListSubItems(10).Text = rPajak!KD_KLS_TANAH
            vOP.ListItems.Item(i).ListSubItems(11).Text = rPajak!NILAI_MIN_TANAH * 1000
            vOP.ListItems.Item(i).ListSubItems(12).Text = rPajak!NILAI_MAX_TANAH * 1000
            vOP.ListItems.Item(i).ListSubItems(13).Text = rPajak!NILAI_PER_M2_TANAH * 1000
        End If
    Next
rPajak.MoveNext
Loop

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub SIMPAN_KLAS1()
On Error GoTo Salah
D_STR = "DELETE  FROM TEMP_BUMI"
openDB (D_STR)
xRec = vOP.ListItems.Count
For i = 1 To vOP.ListItems.Count
Me.Caption = "Persiapan Pencetakan Klasifikasi Tanah : " & Round(i / xRec * 100, 0) & "%"
xKode1 = vOP.ListItems.Item(i).ListSubItems(2).Text
xKec = vOP.ListItems.Item(i).ListSubItems(3).Text
xKode2 = vOP.ListItems.Item(i).ListSubItems(4).Text
xKel = vOP.ListItems.Item(i).ListSubItems(5).Text
xJALAN = vOP.ListItems.Item(i).ListSubItems(6).Text
xZNT = vOP.ListItems.Item(i).ListSubItems(8).Text
xTahun = vOP.ListItems.Item(i).ListSubItems(9).Text
xKelas = vOP.ListItems.Item(i).ListSubItems(10).Text
xMIN = vOP.ListItems.Item(i).ListSubItems(11).Text
xMAX = vOP.ListItems.Item(i).ListSubItems(12).Text
xNJOP = vOP.ListItems.Item(i).ListSubItems(13).Text
xBLOK = vOP.ListItems.Item(i).ListSubItems(14).Text
C_STR = "INSERT INTO TEMP_BUMI (KD_KECAMATAN,NM_KECAMATAN,KD_KELURAHAN, NM_KELURAHAN,[BLOK],NM_JALAN,KD_ZNT,THN_NIR,KLS_TANAH,NILAI_MIN,NILAI_MAX,NJOP) VALUES ('" & xKode1 & "','" & xKec & "','" & xKode2 & "','" & xKel & "','" & xBLOK & "','" & xJALAN & "','" & xZNT & "','" & xTahun & "','" & xKelas & "','" & xMIN & "','" & xMAX & "','" & xNJOP & "')"
openDB (C_STR)
Next

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

