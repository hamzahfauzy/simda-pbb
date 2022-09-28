VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSPPT_Tunggal 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penetapan SPPT Tunggal"
   ClientHeight    =   6930
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   5940
   ControlBox      =   0   'False
   Icon            =   "frmSPPT_Tunggal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   5940
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      Height          =   510
      Left            =   -30
      TabIndex        =   43
      Top             =   -105
      Width           =   5985
      Begin MSMask.MaskEdBox aNOP 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   150
         Width           =   4065
         _ExtentX        =   7170
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
         Left            =   1440
         TabIndex        =   0
         Top             =   150
         Width           =   4005
      End
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
         Left            =   5520
         TabIndex        =   32
         Top             =   165
         Width           =   345
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
         TabIndex        =   44
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   2490
      Left            =   -30
      TabIndex        =   39
      Top             =   315
      Width           =   5970
      Begin VB.TextBox tID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   4425
      End
      Begin VB.TextBox tNOP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2100
         Width           =   4440
      End
      Begin VB.TextBox tNOP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1770
         Width           =   1065
      End
      Begin VB.TextBox tNOP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1770
         Width           =   3045
      End
      Begin VB.TextBox tNOP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         Index           =   6
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   4440
      End
      Begin VB.TextBox tNOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         Left            =   4845
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1110
         Width           =   1020
      End
      Begin VB.TextBox tNOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1110
         Width           =   930
      End
      Begin VB.TextBox tNOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1110
         Width           =   900
      End
      Begin VB.TextBox tNOP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   780
         Width           =   4425
      End
      Begin VB.TextBox tNOP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   450
         Width           =   4425
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "No. ID"
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
         TabIndex        =   60
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
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
         Left            =   120
         TabIndex        =   56
         Top             =   2145
         Width           =   1215
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pos"
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
         Left            =   4500
         TabIndex        =   55
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
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
         Left            =   120
         TabIndex        =   54
         Top             =   1785
         Width           =   1215
      End
      Begin VB.Label Label18 
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
         Left            =   120
         TabIndex        =   53
         Top             =   1455
         Width           =   1215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   4455
         TabIndex        =   52
         Top             =   1155
         Width           =   195
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   2940
         TabIndex        =   51
         Top             =   1155
         Width           =   255
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blok"
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
         Left            =   1440
         TabIndex        =   42
         Top             =   1155
         Width           =   285
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
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
         Height          =   210
         Left            =   105
         TabIndex        =   41
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama WP"
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
         TabIndex        =   40
         Top             =   495
         Width           =   1215
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
      Left            =   3390
      TabIndex        =   31
      Top             =   6375
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
      Left            =   2490
      TabIndex        =   30
      Top             =   6375
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
      Left            =   1590
      TabIndex        =   29
      Top             =   6375
      Width           =   915
   End
   Begin VB.Frame Frame3 
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
      Height          =   2910
      Left            =   -30
      TabIndex        =   33
      Top             =   2685
      Width           =   5985
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
         Index           =   22
         Left            =   1425
         TabIndex        =   24
         Top             =   2175
         Width           =   675
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
         Index           =   21
         Left            =   3990
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0"
         Top             =   2175
         Width           =   1890
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
         Index           =   17
         Left            =   1425
         TabIndex        =   22
         Top             =   1845
         Width           =   1710
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
         Index           =   18
         Left            =   3990
         TabIndex        =   23
         Top             =   1845
         Width           =   1890
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
         Index           =   10
         Left            =   3990
         TabIndex        =   13
         Top             =   180
         Width           =   1890
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
         Index           =   20
         Left            =   3990
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         Top             =   2520
         Width           =   1890
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
         Index           =   19
         Left            =   1425
         TabIndex        =   26
         Text            =   "0"
         Top             =   2505
         Width           =   1710
      End
      Begin VB.TextBox tNOP 
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
         Index           =   13
         Left            =   1425
         TabIndex        =   18
         Top             =   1185
         Width           =   1710
      End
      Begin VB.TextBox tNOP 
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
         Index           =   14
         Left            =   3990
         TabIndex        =   19
         Top             =   1185
         Width           =   1890
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
         Index           =   15
         Left            =   1425
         TabIndex        =   20
         Top             =   1515
         Width           =   1710
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
         Index           =   16
         Left            =   3990
         TabIndex        =   21
         Top             =   1515
         Width           =   1890
      End
      Begin VB.TextBox tNOP 
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
         Index           =   12
         Left            =   3990
         TabIndex        =   17
         Top             =   855
         Width           =   1890
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
         Left            =   1425
         TabIndex        =   12
         Top             =   180
         Width           =   1755
      End
      Begin VB.TextBox tNOP 
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
         Index           =   11
         Left            =   1440
         TabIndex        =   16
         Top             =   855
         Width           =   1710
      End
      Begin MSComCtl2.DTPicker dTerbit 
         Height          =   315
         Left            =   3990
         TabIndex        =   15
         Top             =   525
         Width           =   1920
         _ExtentX        =   3387
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
         Format          =   187105281
         CurrentDate     =   41486
      End
      Begin MSComCtl2.DTPicker dJatuh 
         Height          =   315
         Left            =   1425
         TabIndex        =   14
         Top             =   525
         Width           =   1740
         _ExtentX        =   3069
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
         Format          =   187105281
         CurrentDate     =   41486
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   2145
         TabIndex        =   63
         Top             =   2220
         Width           =   165
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "NJKP"
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
         Left            =   3225
         TabIndex        =   62
         Top             =   2205
         Width           =   600
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faktor Pengurang"
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
         Left            =   105
         TabIndex        =   61
         Top             =   2535
         Width           =   1290
      End
      Begin VB.Label Label24 
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
         Height          =   195
         Left            =   75
         TabIndex        =   59
         Top             =   1875
         Width           =   1305
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "NJOPTKP"
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
         Left            =   3210
         TabIndex        =   58
         Top             =   1875
         Width           =   1305
      End
      Begin VB.Label Label22 
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
         Height          =   210
         Left            =   3225
         TabIndex        =   57
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "NJOP BNG"
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
         Left            =   3210
         TabIndex        =   50
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "NJOP Bumi"
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
         Left            =   75
         TabIndex        =   49
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas BNG"
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
         Left            =   3210
         TabIndex        =   48
         Top             =   1230
         Width           =   1305
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas Bumi"
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
         Left            =   75
         TabIndex        =   47
         Top             =   1215
         Width           =   1305
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Luas BNG"
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
         Left            =   3225
         TabIndex        =   46
         Top             =   915
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Luas Bumi"
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
         Left            =   75
         TabIndex        =   45
         Top             =   885
         Width           =   1305
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PBB"
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
         Left            =   3255
         TabIndex        =   38
         Top             =   2550
         Width           =   1305
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
         Left            =   3195
         TabIndex        =   37
         Top             =   555
         Width           =   1215
      End
      Begin VB.Label Label8 
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
         Height          =   210
         Left            =   75
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarif"
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
         Left            =   90
         TabIndex        =   35
         Top             =   2220
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
         Left            =   105
         TabIndex        =   34
         Top             =   570
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   30
      TabIndex        =   64
      Top             =   5565
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
         TabIndex        =   28
         Top             =   180
         Width           =   4260
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
         TabIndex        =   65
         Top             =   225
         Width           =   1305
      End
   End
   Begin VB.Image Image1 
      Height          =   9180
      Left            =   0
      Picture         =   "frmSPPT_Tunggal.frx":1CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "frmSPPT_Tunggal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xTT, xTB
Dim NMIN, cTarif
Dim xMIN(2), xMAX(2)
Dim xTarif(2)
Dim cMin(2), cMax(2), cTKP(2)
Dim totChar
Dim PBB_Bayar
Private Sub cmdBangunan_Click()
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
frmOP_Tanah.Show
End Sub

Private Sub aNOP_LostFocus()
On Error Resume Next
'If chPajak(1).Value = 1 Or chPajak(2).Value = 1 Then
    'Panggil Data Bangunan
    'StrQ1 = "Select * From QOBJEKPAJAK WHERE NOPQ =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
    'openDB (StrQ1)
    call_data
    tNOP(0).Text = aNOP.Text
'Else
'    StrQ1 = "Select * From DAT_OP_BUMI WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_URUT ASC"
'    openDB (StrQ1)
'    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    If rPajak.EOF Then
'        MsgBox "Nomor Objerk Pajak tidak terdaftar...", vbCritical, "Tetnong...!"
'    End If
'End If

End Sub

Private Sub ccBayar_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
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
    KeyAscii = 0
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

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

Private Sub cmdClear_Click()
On Error Resume Next
For Each Control In Me
If TypeOf Control Is TextBox Then
    Control.Text = 0
ElseIf TypeOf Control Is ComboBox Then
    Control.Text = ""
End If
Next
tID.Text = ""
For i = 1 To 9
    tNOP(i).Text = ""
Next
tNOP(3).Text = "00"
tNOP(4).Text = "00"
tNOP(5).Text = "00"
tNOP(10).Text = "00"
dJatuh.Value = Format(Now, "dd/mm/yyyy")
dTerbit.Value = Format(Now, "dd/mm/yyyy")
ccTahun.Text = ccTahun.List(0)
aNOP.SetFocus
ccBayar.Text = ccBayar.List(3)
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
    xID = 4
    frmLIST_Objek1.Show
End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo Salah
If (Trim(tNOP(11).Text) = "" Or Trim(tNOP(11).Text) = 0) And (Trim(tNOP(12).Text) = "" Or Trim(tNOP(12).Text) = 0) Then
    MsgBox "Luas Bumi dan Bangunan : 0, proses tidak dilanjutkan", vbCritical, "Error"
    Exit Sub
End If
xSTR = "Select THN_NJOPTKP From DAT_SUBJEK_PAJAK_NJOPTKP WHERE THN_NJOPTKP='" & ccTahun.Text & "' ORDER BY THN_NJOPTKP ASC"
openDB (xSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If rPajak.EOF Then
    MsgBox "NJOPTKP Untuk tahun " & ccTahun.Text & " Beluma dibuat" & _
            vbCrLf & "Proses tidak akan dilanjutkan...", vbCritical, "Tetnong"
    Exit Sub
End If
strLOG = "SELECT * FROM LOGUTAMA WHERE NOP1='" & Trim(aNOP.Text) & "' ORDER BY TGL_PEREKAMAN_OP ASC"

'strLOG = "iLOG '" & ccTahun.Text * 1 & "'"
openDB (strLOG)
Do While Not rPajak.EOF
    rPajak!KD_KLS_TANAH = tNOP(13).Text
    rPajak!KD_KLS_BNG = tNOP(14).Text
    rPajak!NJOP_BUMI1 = tNOP(15).Text
    rPajak!NJOP_BNG1 = tNOP(16).Text
    rPajak!NJOPTKP1 = tNOP(18).Text
    rPajak!PBB_TERUTANG1 = tNOP(20).Text
    rPajak.Update
rPajak.MoveNext
Loop
sv_SPPT



If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub dJatuh_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
End If

End Sub

Private Sub dTerbit_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
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
dJatuh.Value = Format(Now, "dd/mm/yyyy")
dTerbit.Value = Format(Now, "dd/mm/yyyy")

XXSTR1 = "select  THN_AWAL_KLS_TANAH  from Kelas_tanah order by THn_AWAL_KLS_TANAH DESC"
openDB (XXSTR1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then xTT = rPajak!THN_AWAL_KLS_TANAH
XXSTR2 = "select  THN_AWAL_KLS_BNG from KELAS_BANGUNAN order by THN_AWAL_KLS_BNG DESC"
openDB (XXSTR2)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then xTB = rPajak!THN_AWAL_KLS_BNG
CALL_TBAYAR
ccBayar.Text = ccBayar.List(3)
'MsgBox xID
'If xID = 4 Then
'MsgBox XTT
'    CALL_TARIF
'    If tNOP(17).Text >= xMIN(1) And tNOP(17).Text <= xMAX(1) Then
'        cTarif = xTarif(1)
'    Else
'        cTarif = xTarif(2)
'    End If
'    K_BUMI
'    K_BANGUNAN
'    tNOP(21).Text = Format(tNOP(17).Text * 1 - tNOP(18).Text * 1 - tNOP(19).Text * 1, "#,#0")
'    tNOP(20).Text = Format((tNOP(21).Text * cTarif / 100), "#,#0")
'
'    Call_MIN
'    tNOP(22).Text = cTarif & " %"
'    If tNOP(21).Text < 0 Then tNOP(21).Text = 0
'    If tNOP(20).Text < NMIN Then
'        tNOP(20).Text = Format(NMIN, "#,#0")
'    End If
'
'End If
'CEK_JLH
End Sub
Sub K_BUMI()
On Error GoTo Salah
xxSTR = "select * from Kelas_TANAH WHERE THN_AWAL_KLS_TANAH='" & xTT & "'"
openDB (xxSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
L_BUMI = tNOP(11).Text
If L_BUMI <= 0 Then L_BUMI = 1
Do While Not rPajak.EOF
    If Format(tNOP(15).Text / L_BUMI, "#,#0") = Format(rPajak!NILAI_PER_M2_TANAH * 1000, "#,#0") Then
        tNOP(13).Text = rPajak!KD_KLS_TANAH
        Exit Sub
    End If
rPajak.MoveNext
Loop
tNOP(13).Text = 0
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
Sub K_BANGUNAN()
'On Error Resume Next
On Error GoTo Salah
xxSTR = "select * from Kelas_BANGUNAN WHERE THN_AWAL_KLS_BNG='" & xTB & "'"
openDB (xxSTR)
L_BNG = tNOP(12).Text
If L_BNG <= 0 Then L_BNG = 1
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    If Format(tNOP(16).Text / L_BNG, "#,#0") = Format(rPajak!NILAI_PER_M2_BNG * 1000, "#,#0") Then
        tNOP(14).Text = rPajak!KD_KLS_BNG
        Exit Sub
        
    End If
rPajak.MoveNext
Loop
tNOP(14).Text = 0
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub CALL_TARIF()
On Error GoTo Salah
xxSTR = "Select * From Tarif order by NJOP_MIN"
openDB (xxSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
Do While Not rPajak.EOF
    i = i + 1
    xMIN(i) = rPajak!NJOP_MIN
    xMAX(i) = rPajak!NJOP_MAX
    xTarif(i) = rPajak!NILAI_TARIF
rPajak.MoveNext
Loop
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
Do While Not rPajak.EOF
    NMIN = rPajak!NILAI_PBB_MINIMAL
rPajak.MoveNext
Loop
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
    KeyAscii = 0
End If
End Sub
Sub call_data()
On Error GoTo Salah
Screen.MousePointer = vbHourglass

    StrQ1 = "Select * From QOBJEKPAJAK WHERE NOPQ =  '" & Trim(aNOP.Text) & "' ORDER BY nopq asc"
    openDB (StrQ1)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        MsgBox "Data Tidak Ditemukan...", vbCritical, "Tetnong..."
        Screen.MousePointer = vbDefault
        Exit Sub
    'End If
    'Do While Not rPajak.EOF
    Else
     If rPajak!JNS_BUMI = "4" Then MsgBox "SPPT Untuk Fasilitas Umum tidak perlu", vbExclamation, "Tetnong": Screen.MousePointer = vbDefault: Exit Sub
     tNOP(0).Text = rPajak!NOPQ
     aNOP.Text = rPajak!NOPQ
     tID.Text = Trim(rPajak!SUBJEK_PAJAK_ID)
    tNOP(1).Text = rPajak!Nm_wp
    tNOP(11).Text = Format(rPajak!TOTAL_LUAS_BUMI, "#,#0")
    tNOP(12).Text = Format(rPajak!TOTAL_LUAS_BNG, "#,#0")
    tNOP(15).Text = Format(rPajak!NJOP_BUMI, "#,#0")
    tNOP(16).Text = Format(rPajak!NJOP_BNG, "#,#0")
    tNOP(2).Text = rPajak!JALAN_WP
    If IsNull(rPajak!BLOK_KAV_NO_WP) = True Or rPajak!BLOK_KAV_NO_WP = "" Then rPajak!BLOK_KAV_NO_WP = "00"
    tNOP(3).Text = rPajak!BLOK_KAV_NO_WP
    If IsNull(rPajak!RW_WP) = True Or rPajak!RW_WP = "" Then rPajak!RW_WP = "00"
    tNOP(4).Text = rPajak!RW_WP
    If IsNull(rPajak!RT_WP) = True Or rPajak!RT_WP = "" Then rPajak!RT_WP = "00"
    tNOP(5).Text = rPajak!RT_WP
    If IsNull(rPajak!KELURAHAN_WP) = True Or rPajak!KELURAHAN_WP = "" Then rPajak!KELURAHAN_WP = "-"
    tNOP(6).Text = rPajak!KELURAHAN_WP
    If IsNull(rPajak!KOTA_WP) = True Or rPajak!KOTA_WP = "" Then rPajak!KOTA_WP = "-"
    tNOP(7).Text = rPajak!KOTA_WP
    If IsNull(rPajak!KD_POS_WP) = True Or rPajak!KD_POS_WP = "" Then rPajak!KD_POS_WP = "00000"
    tNOP(8).Text = rPajak!KD_POS_WP
    If IsNull(rPajak!NPWP) = True Or rPajak!NPWP = "" Then rPajak!NPWP = "-"
    tNOP(9).Text = rPajak!NPWP
   '
    If IsNull(rPajak!NO_PERSIL) = True Or rPajak!NO_PERSIL = "" Then rPajak!NO_PERSIL = "00"
    tNOP(10).Text = rPajak!NO_PERSIL
    'rPajak.MoveNext
    'Loop
    End If
    Call_Proses
    
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
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
Do While Not rPajak.EOF
    i = i + 1
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

Sub J_Karakter()
On Error Resume Next
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
End Sub

Private Sub tID_GotFocus()
Call c_blok(tID)
End Sub

Private Sub tID_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If

End Sub

Private Sub tNOP_GotFocus(Index As Integer)
'Select Case Index
'Case 3, 4, 5, 10, 11 To 21
    Call c_blok(tNOP(Index))
'End Select
End Sub

Private Sub tNOP_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
Select Case Index
Case 3, 4, 5, 10, 11 To 22
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select
End Sub

Private Sub tNOP_LostFocus(Index As Integer)
On Error GoTo Salah
Select Case Index
Case 3, 4, 5, 10, 11 To 22
    Call c_Kosong(tNOP(Index))
End Select
Select Case Index
Case 18, 19, 22
    'tNOP(20).Text = Format((tNOP(21).Text * cTarif / 100) - tNOP(19).Text * 1, "#,#0")
    tNOP(21).Text = tNOP(17).Text * 1 - tNOP(18).Text * 1
    If tNOP(21).Text <= 0 Then tNOP(21).Text = 0
    tNOP(20).Text = Format((tNOP(21).Text * tNOP(22).Text * 1 / 100) - tNOP(19).Text * 1, "#,#0")
    Call_MIN
    If tNOP(20).Text * 1 < NMIN * 1 Then
        tNOP(20).Text = Format(NMIN, "#,#0")
    End If
    
End Select
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
'----Prosedur Menentukan NJOPTKP Individual, dengan ketentuan : Hanya dikenakan untuk jenis bumi = 1 dan hanya untuk satu NOP saja
Sub CEK_NJOPTKP()
On Error GoTo Salah
'List1.Clear
n_STR = "select * from DAT_SUBJEK_PAJAK_NJOPTKP WHERE (SUBJEK_PAJAK_ID)='" & Trim(tID.Text) & "' ORDER BY SUBJEK_PAJAK_ID ASC"
'n_STR = "select * from DAT_SUBJEK_PAJAK_NJOPTKP ORDER BY SUBJEK_PAJAK_ID ASC"
openDB (n_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'rPajak.Find " SUBJEK_PAJAK_ID='" & Trim(tID.Text) & "' "
If rPajak.EOF Then
    tNOP(18).Text = 0
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
Sub c_Kosong(nControl As TextBox)
On Error Resume Next
If nControl.Text = "" Or nControl.Text = "-" Or nControl.Text = "." Then
    nControl.Text = 0
End If
nControl.Alignment = 1
End Sub

Private Sub vBangunan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan.SortKey = ColumnHeader.Index - 1
vBangunan.Sorted = True
vBangunan.Sorted = False
vBangunan.SortOrder = lvwAscending
End Sub
Sub Call_Proses()
On Error GoTo Salah
tNOP(17).Text = Format(tNOP(15).Text * 1 + tNOP(16).Text * 1, "#,#0")
    CALL_NJOPTKP
    If tNOP(17).Text * 1 >= cMin(1) And tNOP(17).Text * 1 <= cMax(1) Then
        tNOP(18).Text = Format(cTKP(1), "#,#0")
        cTarif = xTarif(1)
    Else
        tNOP(18).Text = Format(cTKP(2), "#,#0")
        cTarif = xTarif(2)
    End If
    'If tNOP(12).Text * 1 <= 0 Or tNOP(16).Text * 1 <= 0 Then tNOP(18).Text = 0
    CEK_NJOPTKP
    tNOP(22).Text = cTarif
    K_BUMI
    K_BANGUNAN
    tNOP(21).Text = Format(tNOP(17).Text * 1 - tNOP(18).Text * 1, "#,#0")
    
    'Nilai Sebelum Dikurangi
    tNOP(20).Text = Format(tNOP(21).Text * tNOP(22).Text * 1 / 100, "#,#0") 'PBB Terutang
    'Nilai Setelah Dikurangi
        PBB_Bayar = Format((tNOP(20).Text * 1) - (tNOP(19).Text * 1), "#,#0")
    Call_MIN
    
    If tNOP(21).Text * 1 < 0 Then tNOP(21).Text = 0
    If tNOP(20).Text * 1 < NMIN * 1 Then
        tNOP(20).Text = Format(NMIN, "#,#0")
    End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub sv_SPPT()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
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
For i = 1 To 22
    If tNOP(i).Text = "" Or tID.Text = "" Then
        MsgBox "Masih ada data kosong...", vbCritical, "Tetnong..."
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
Next
'Cek Keberadaan NOP
StrQ1 = "Select * From QOBJEKPAJAK WHERE NOPQ =  '" & Trim(aNOP.Text) & "' ORDER BY nopq asc"
    openDB (StrQ1)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        MsgBox "Data Tidak Ditemukan...", vbCritical, "Tetnong..."
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    xxKec = Mid(Trim(aNOP.Text), 7, 3)
xxKel = Mid(Trim(aNOP.Text), 11, 3)
xxBlok = Mid(Trim(aNOP.Text), 15, 3)
xxUrut = Mid(Trim(aNOP.Text), 19, 4)
xxJenis = Right(Trim(aNOP.Text), 1)
xHutang = Round(tNOP(22).Text / 100 * tNOP(21).Text, 0)
xKurang = Round(tNOP(19).Text * 1, 0)
'Format((tNOP(20).Text * 1) - (tNOP(19).Text * 1), "#,#0")
xBayar = Round((tNOP(20).Text * 1) - (tNOP(19).Text * 1), 0) 'Format(tNOP(20).Text * 1 - PBB_Bayar * 1, "#,#0")
xxTerbit = Format(dTerbit.Value, "yyyy-mm-dd")
xxJTempo = Format(dJatuh.Value, "yyyy-mm-dd")

xSQL = "Select * From SPPT where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & Trim(aNOP.Text) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'"

openDB (xSQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then
        If rPajak!PROSES = "T" Or rPajak!PROSES = "M" Then
            TANYA = MsgBox("SPPT dengan NOP : " & aNOP.Text & " sudah ada..." & _
                vbCrLf & "Hapus yang sudah ada?", vbCritical + vbYesNo, "Tetnong")
                CPP = rPajak!PROSES
        Else
            TANYA = MsgBox("Objek Pajak Sudah Dinilai, Lanjutkan?", vbQuestion + vbYesNo, "Tetnong")
            CPP = rPajak!PROSES
        End If
        If TANYA = vbYes Then
            'xSQL = "Delete From SPPT where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & Trim(aNOP.Text) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "' AND PROSES='" & CPP & "'"
            'iSQL = "iSPPT_TUNGGAL '12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '" & ccTahun.Text & "', '" & tNOP(1).Text & "', '" & tNOP(2).Text & "', '" & tNOP(3).Text & "', '" & tNOP(4).Text & "', '" & tNOP(5).Text & "','" & tNOP(6).Text & "','" & tNOP(7).Text & "','" & tNOP(8).Text & "','" & tNOP(9).Text & "','" & tNOP(10).Text & "'," & _
            " '" & tNOP(13).Text & " ','" & xTT & "','" & tNOP(14).Text & "','" & xTB & "','" & xxJTempo & "','" & tNOP(11).Text & "','" & tNOP(12).Text & "','" & tNOP(15).Text & "','" & tNOP(16).Text & "','" & tNOP(17).Text & "','" & tNOP(18).Text & "','" & tNOP(21).Text & "','" & xHutang & "','" & xKurang & "','" & xBayar & "','0','0','0','" & xxTerbit & "','" & xxTerbit & "','000000', " & _
            " 1,'01','16','04','01','" & Left(Trim(ccBayar.Text), 2) & "','T','" & aNOP.Text & "'"
            'openDB (iSQL)
            GoTo Keluar
            
        Else
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
Else
       GoTo Keluar
End If

Keluar:
'iSQL = "iSPPT_TUNGGAL  '12', '12','" & xxKec & "','" & xxKel & "','" & xxBlok & "','" & xxUrut & "','" & xxJenis & "','" & ccTahun.Text & "','" & tNOP(1).Text & "','" & tNOP(2).Text & "','" & tNOP(3).Text & "','" & tNOP(4).Text & "','" & tNOP(5).Text & "'," & _
        "'" & tNOP(6).Text & "','" & tNOP(7).Text & "','" & tNOP(8).Text * 1 & "','" & tNOP(9).Text & "','" & tNOP(10).Text & "','" & tNOP(13).Text & "','" & xTT & "','" & tNOP(14).Text & "','" & xTB & "','" & xxJTempo & "','" & Round(tNOP(11).Text, 0) & "','" & Round(tNOP(12).Text, 0) & "','" & Round(tNOP(15).Text, 0) & "'," & _
        "'" & Round(tNOP(16).Text, 0) & "','" & Round(tNOP(17).Text, 0) & "','" & Round(tNOP(18).Text, 0) & "','" & Round(tNOP(21).Text, 0) & "','" & Round(xHutang, 0) & "','" & Round(xKurang, 0) & "','" & Round(xBayar, 0) & "', '0', '0', '0','" & xxTerbit & "','" & xxTerbit & "', '000000',1, '01', '16', '04', '01','" & Left(Trim(ccBayar.Text), 2) & "', 'T','" & CPP & "','" & aNOP.Text & "'"
 '       openDB (iSQL)


'xxKec = Mid(Trim(aNOP.Text), 7, 3)
'xxKel = Mid(Trim(aNOP.Text), 11, 3)
'xxBlok = Mid(Trim(aNOP.Text), 15, 3)
'xxUrut = Mid(Trim(aNOP.Text), 19, 4)
'xxJenis = Right(Trim(aNOP.Text), 1)
'xHutang = Format(tNOP(22).Text / 100 * tNOP(21).Text, "#,#0")
'xKurang = Format(tNOP(19).Text * 1, "#,#0")
''Format((tNOP(20).Text * 1) - (tNOP(19).Text * 1), "#,#0")
'xBayar = Format((tNOP(20).Text * 1) - (tNOP(19).Text * 1), "#,#0") 'Format(tNOP(20).Text * 1 - PBB_Bayar * 1, "#,#0")
'xxTerbit = Format(dTerbit.Value, "dd/mm/yyyy")
'xxJTempo = Format(dJatuh.Value, "dd/mm/yyyy")
''iSQL = "INSERT INTO SPPT(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,THN_PAJAK_SPPT,NM_WP_SPPT,JLN_WP_SPPT,BLOK_KAV_NO_WP_SPPT,RW_WP_SPPT,RT_WP_SPPT,KELURAHAN_WP_SPPT,KOTA_WP_SPPT,KD_POS_WP_SPPT,NPWP_SPPT," & _
'    "NO_PERSIL_SPPT,KD_KLS_TANAH,THN_AWAL_KLS_TANAH,KD_KLS_BNG,THN_AWAL_KLS_BNG,TGL_JATUH_TEMPO_SPPT,LUAS_BUMI_SPPT,LUAS_BNG_SPPT, NJOP_BUMI_SPPT,NJOP_BNG_SPPT,NJOP_SPPT,NJOPTKP_SPPT,NJKP_SPPT,PBB_TERHUTANG_SPPT,FAKTOR_PENGURANG_SPPT," & _
'    "PBB_YG_HARUS_DIBAYAR_SPPT,STATUS_PEMBAYARAN_SPPT,STATUS_TAGIHAN_SPPT,STATUS_CETAK_SPPT,TGL_TERBIT_SPPT,TGL_CETAK_SPPT,NIP_PENCETAK_SPPT,SIKLUS_SPPT,KD_KANWIL_BANK,KD_KPPBB_BANK,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,KD_TP,PROSES)" & _
'    "Values('12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '" & ccTahun.Text & "', '" & tNOP(1).Text & "', '" & tNOP(2).Text & "', '" & tNOP(3).Text & "', '" & tNOP(4).Text & "', '" & tNOP(5).Text & "','" & tNOP(6).Text & "','" & tNOP(7).Text & "','" & tNOP(8).Text & "','" & tNOP(9).Text & "','" & tNOP(10).Text & "'," & _
'    " '" & tNOP(13).Text & " ','" & xTT & "','" & tNOP(14).Text & "','" & xTB & "','" & xxJTempo & "','" & tNOP(11).Text & "','" & tNOP(12).Text & "','" & tNOP(15).Text & "','" & tNOP(16).Text & "','" & tNOP(17).Text & "','" & tNOP(18).Text & "','" & tNOP(21).Text & "','" & xHutang & "','" & xKurang & "','" & xBayar & "','0','0','0','" & xxTerbit & "','" & xxTerbit & "','000000', " & _
'    " 1,'01','16','04','01','01','T')"
' '   openDB (iSQL)
xSQL = "Delete From SPPT where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & Trim(aNOP.Text) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "' AND PROSES='" & CPP & "'"
openDB (xSQL)
iSQL = "Select * From SPPT where ('12.12.' + KD_KECAMATAN + '.' + KD_KELURAHAN+ '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP)='" & Trim(aNOP.Text) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "' AND PROSES='" & CPP & "'"
openDB (iSQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
rPajak.AddNew
rPajak!KD_PROPINSI = "12"
rPajak!KD_DATI2 = "12"
rPajak!KD_KECAMATAN = xxKec
rPajak!KD_KELURAHAN = xxKel
rPajak!KD_BLOK = xxBlok
rPajak!NO_URUT = xxUrut
rPajak!KD_JNS_OP = xxJenis
rPajak!THN_PAJAK_SPPT = ccTahun.Text
rPajak!NM_WP_SPPT = tNOP(1).Text
rPajak!JLN_WP_SPPT = tNOP(2).Text
rPajak!BLOK_KAV_NO_WP_SPPT = tNOP(3).Text
rPajak!RW_WP_SPPT = tNOP(4).Text
rPajak!RT_WP_SPPT = tNOP(5).Text
rPajak!KELURAHAN_WP_SPPT = tNOP(6).Text
rPajak!KOTA_WP_SPPT = tNOP(7).Text
rPajak!KD_POS_WP_SPPT = tNOP(8).Text
rPajak!NPWP_SPPT = tNOP(9).Text
rPajak!NO_PERSIL_SPPT = tNOP(10).Text
rPajak!KD_KLS_TANAH = tNOP(13).Text
rPajak!THN_AWAL_KLS_TANAH = xTT
rPajak!KD_KLS_BNG = tNOP(14).Text
rPajak!THN_AWAL_KLS_BNG = xTB
rPajak!TGL_JATUH_TEMPO_SPPT = xxJTempo
rPajak!LUAS_BUMI_SPPT = tNOP(11).Text
rPajak!LUAS_BNG_SPPT = tNOP(12).Text
rPajak!NJOP_BUMI_SPPT = tNOP(15).Text
rPajak!NJOP_BNG_SPPT = tNOP(16).Text
rPajak!NJOP_SPPT = tNOP(17).Text
rPajak!NJOPTKP_SPPT = tNOP(18).Text
rPajak!NJKP_SPPT = tNOP(21).Text
rPajak!PBB_TERHUTANG_SPPT = xHutang
rPajak!FAKTOR_PENGURANG_SPPT = xKurang
rPajak!PBB_YG_HARUS_DIBAYAR_SPPT = xBayar
rPajak!STATUS_PEMBAYARAN_SPPT = "0"
rPajak!STATUS_TAGIHAN_SPPT = "0"
rPajak!STATUS_CETAK_SPPT = "0"
rPajak!TGL_TERBIT_SPPT = xxTerbit
rPajak!TGL_CETAK_SPPT = xxTerbit
rPajak!NIP_PENCETAK_SPPT = "000000"
rPajak!SIKLUS_SPPT = 1
rPajak!KD_KANWIL_BANK = "01"
rPajak!KD_KPPBB_BANK = "16"
rPajak!KD_BANK_TUNGGAL = "04"
rPajak!KD_BANK_PERSEPSI = "01"
rPajak!KD_TP = Left(Trim(ccBayar.Text), 2)
rPajak!PROSES = "T"

rPajak.Update

If Err.Number = 0 Then
    MsgBox "SPPT Berhasil dibuat....!", vbInformation, "Sukses!"
    cmdClear_Click
Else
    MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error"
End If
'rPajak.AddNew
'rPajak!KD_PROPINSI = "12"
'rPajak!KD_DATI2 = "12"
'rPajak!KD_KECAMATAN = xxKec
'rPajak!KD_KELURAHAN = xxKel
'rPajak!KD_BLOK = xxBlok
'rPajak!NO_URUT = xxUrut
'rPajak!KD_JNS_OP = xxJenis
'rPajak!THN_PAJAK_SPPT = ccTahun.Text
'rPajak!NM_WP_SPPT = tNOP(1).Text
'rPajak!JLN_WP_SPPT = tNOP(2).Text
'rPajak!BLOK_KAV_NO_WP_SPPT = tNOP(3).Text
'rPajak!RW_WP_SPPT = tNOP(4).Text
'rPajak!RT_WP_SPPT = tNOP(5).Text
'rPajak!KELURAHAN_WP_SPPT = tNOP(6).Text
'rPajak!KOTA_WP_SPPT = tNOP(7).Text
'rPajak!KD_POS_WP_SPPT = tNOP(8).Text
'rPajak!NPWP_SPPT = tNOP(9).Text
'rPajak!NO_PERSIL_SPPT = tNOP(10).Text
'rPajak!KD_KLS_TANAH = tNOP(13).Text
'rPajak!THN_AWAL_KLS_TANAH = xTT
'rPajak!KD_KLS_BNG = tNOP(14).Text
'rPajak!THN_AWAL_KLS_BNG = xTB
'rPajak!TGL_JATUH_TEMPO_SPPT = Format(dJatuh.Value, "DD/MM/YYYY")
'rPajak!LUAS_BUMI_SPPT = tNOP(11).Text
'rPajak!LUAS_BNG_SPPT = tNOP(12).Text
'rPajak!NJOP_BUMI_SPPT = tNOP(15).Text
'rPajak!NJOP_BNG_SPPT = tNOP(16).Text
'rPajak!NJOP_SPPT = tNOP(17).Text
'rPajak!NJOPTKP_SPPT = tNOP(18).Text
'rPajak!NJKP_SPPT = tNOP(21).Text
'rPajak!PBB_TERHUTANG_SPPT = xHutang
'rPajak!FAKTOR_PENGURANG_SPPT = xKurang
'rPajak!PBB_YG_HARUS_DIBAYAR_SPPT = xBayar
'rPajak!STATUS_PEMBAYARAN_SPPT = "0"
'rPajak!STATUS_TAGIHAN_SPPT = "0"
'rPajak!STATUS_CETAK_SPPT = "0"
'rPajak!TGL_TERBIT_SPPT = xxTerbit
'rPajak!TGL_CETAK_SPPT = xxTerbit
'rPajak!NIP_PENCETAK_SPPT = "000000"
'rPajak!SIKLUS_SPPT = 1
'rPajak!KD_KANWIL_BANK = "01"
'rPajak!KD_KPPBB_BANK = "16"
'rPajak!KD_BANK_TUNGGAL = "04"
'rPajak!KD_BANK_PERSEPSI = "01"
'rPajak!KD_TP = "09"
'rPajak!PROSES = "T"
'rPajak.Update


If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub
