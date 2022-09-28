VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNilai_Massal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penilaian Massal"
   ClientHeight    =   3540
   ClientLeft      =   135
   ClientTop       =   420
   ClientWidth     =   11505
   ControlBox      =   0   'False
   Icon            =   "frmNilai_Massal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   11505
   Begin MSComctlLib.ListView vOP 
      Height          =   3030
      Left            =   6150
      TabIndex        =   21
      Top             =   360
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   5345
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
      NumItems        =   60
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
         Text            =   "ZNT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "NIR LAMA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "NIR BARU"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ket"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "KELAS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "NJOP"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "MIN"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "MAX"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "6"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "7"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "8"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "9"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "10"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Text            =   "11"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   22
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Text            =   "Ket"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "TOTAL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "NJOPTK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Text            =   "sistem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "TIPE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "JPB1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "JPB2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "JPB3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "JPB4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "JPB5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "JPB6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "JPB7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "JPB8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "SUBJEK ID"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "NO FORM"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "JPB11"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "JPB12"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Text            =   "JPB13"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "JPB14"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "JPB15"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "JPB16"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "JPB17"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "L_MEZANINE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "N_MEZANIN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "D_DUKUNG"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "N_D_DUKUNG"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "LBR_BENTANG"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "TG_KOLOM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "KLS_JPB"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "BINTANG"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "JLH_KMR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "L_KMR_AC_SENT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Text            =   "L_R_LAIN_AC_SENT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "Atap"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "Dinding"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "Lantai"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "Langit2"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView vBangunan 
      Height          =   3480
      Left            =   12915
      TabIndex        =   18
      Top             =   15
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   6138
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
      NumItems        =   18
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
         Text            =   "ZNT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "NIR LAMA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "NIR BARU"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ket"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "KELAS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "NJOP"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "MIN"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "MAX"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "6"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "NOP"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Export to Excel"
      Height          =   600
      Left            =   4905
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   375
      Top             =   2535
   End
   Begin VB.CommandButton mnCetak1 
      Caption         =   "&Cetak OP"
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
      Left            =   4695
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSComctlLib.ListView vFAS 
      Height          =   4650
      Left            =   12735
      TabIndex        =   20
      Top             =   3855
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   8202
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
      NumItems        =   60
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
         Text            =   "ZNT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "NIR LAMA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "NIR BARU"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ket"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "KELAS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "NJOP"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "MIN"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "MAX"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "6"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "7"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "8"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "9"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "10"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Text            =   "11"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   22
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Text            =   "Ket"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "TOTAL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "NJOPTK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Text            =   "sistem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "TIPE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "JPB1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "JPB2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "JPB3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "JPB4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "JPB5"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "JPB6"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "JPB7"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "JPB8"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "JPB9"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "JPB10"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "JPB11"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "JPB12"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Text            =   "JPB13"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "JPB14"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "JPB15"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "JPB16"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "JPB17"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "L_MEZANINE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "N_MEZANIN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "D_DUKUNG"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "N_D_DUKUNG"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "LBR_BENTANG"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "TG_KOLOM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "KLS_JPB"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "BINTANG"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "JLH_KMR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "L_KMR_AC_SENT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Text            =   "L_R_LAIN_AC_SENT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "Atap"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "Dinding"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "Lantai"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "Langit2"
         Object.Width           =   2540
      EndProperty
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
      Left            =   3615
      TabIndex        =   9
      Top             =   2730
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
      Left            =   2715
      TabIndex        =   8
      Top             =   2730
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
      Left            =   1815
      TabIndex        =   7
      Top             =   2730
      Width           =   915
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   135
      TabIndex        =   12
      Top             =   75
      Width           =   5745
      Begin VB.CheckBox cPro 
         Caption         =   "Cek Proses"
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
         Left            =   3240
         TabIndex        =   25
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox xPro 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1665
         Width           =   3990
      End
      Begin VB.TextBox xPro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   4065
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   1320
         Width           =   1440
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
         Top             =   870
         Width           =   4005
      End
      Begin VB.TextBox tJum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   1320
         Width           =   1530
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
         Width           =   1530
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
         Width           =   4005
      End
      Begin MSComctlLib.ProgressBar pNilai 
         Height          =   255
         Left            =   1500
         TabIndex        =   23
         Top             =   2220
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   15
         X2              =   5670
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         X1              =   30
         X2              =   5685
         Y1              =   2475
         Y2              =   2475
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   5655
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   45
         X2              =   5700
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "N. O. P"
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
         Left            =   225
         TabIndex        =   22
         Top             =   1710
         Width           =   1305
      End
      Begin VB.Label LNilai 
         BackColor       =   &H8000000B&
         Caption         =   "NOP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1530
         TabIndex        =   6
         Top             =   2025
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Filtering"
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
         Left            =   3315
         TabIndex        =   17
         Top             =   1365
         Width           =   570
      End
      Begin VB.Label Label1 
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
         Left            =   165
         TabIndex        =   16
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Kelurahan/Desa"
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
         TabIndex        =   15
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label9 
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
         Left            =   180
         TabIndex        =   14
         Top             =   555
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Jumlah Record"
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
         TabIndex        =   13
         Top             =   1365
         Width           =   1305
      End
   End
   Begin MSComctlLib.ListView vBng 
      Height          =   4860
      Left            =   75
      TabIndex        =   19
      Top             =   3870
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   8573
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
      NumItems        =   87
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
         Text            =   "ZNT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "NIR LAMA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "NIR BARU"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ket"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "KELAS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "NJOP"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "MIN"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "MAX"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "6"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "7"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "8"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "9"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "10"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Text            =   "11"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   22
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Text            =   "Ket"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "TOTAL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "NJOPTK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Text            =   "sistem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "TIPE"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "JPB1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "JPB2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "JPB3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "JPB4"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "JPB5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "JPB6"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "JPB7"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "JPB8"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "JPB9"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "JPB10"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "JPB11"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "JPB12"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Text            =   "JPB13"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "JPB14"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "JPB15"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "JPB16"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Text            =   "JPB17"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Text            =   "L_MEZANINE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Text            =   "N_MEZANIN"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Text            =   "D_DUKUNG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Text            =   "N_D_DUKUNG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Text            =   "LBR_BENTANG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Text            =   "TG_KOLOM"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Text            =   "KLS_JPB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Text            =   "BINTANG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Text            =   "JLH_KMR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Text            =   "L_KMR_AC_SENT"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Text            =   "L_R_LAIN_AC_SENT"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Text            =   "Atap"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Text            =   "Dinding"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Text            =   "Lantai"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Text            =   "Langit2"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   60
         Text            =   "FASILITAS1"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Text            =   "FAS2"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   62
         Text            =   "Susut"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   63
         Text            =   "Nilai Total"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   64
         Text            =   "NOP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   65
         Text            =   "CC "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   66
         Text            =   "AC_KMR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   67
         Text            =   "AC_LAIN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   68
         Text            =   "BOILER"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   69
         Text            =   "KET"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   70
         Text            =   "Susut_Fas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(72) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   71
         Text            =   "Dindin_DD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(73) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   72
         Text            =   "Formulir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(74) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   73
         Text            =   "J_Transaksi"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(75) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   74
         Text            =   "T_DATA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(76) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   75
         Text            =   "NIP_DATA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(77) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   76
         Text            =   "T_PERIKSA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(78) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   77
         Text            =   "NIP_PERIKSA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(79) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   78
         Text            =   "T_REKAM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(80) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   79
         Text            =   "NIP_REKAM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(81) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   80
         Text            =   "N_INDIVIDU"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(82) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   81
         Text            =   "K_UTAMA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(83) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   82
         Text            =   "K_MATERIAL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(84) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   83
         Text            =   "K_FAS1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(85) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   84
         Text            =   "J_SUSUT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(86) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   85
         Text            =   "K_SUSUT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(87) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   86
         Text            =   "K_FAS2"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "OBJEK PAJAK"
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
      Left            =   8940
      TabIndex        =   24
      Top             =   90
      Width           =   1215
   End
End
Attribute VB_Name = "frmNilai_Massal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nMezanin
Dim xTT, xTB
Dim nSistem_B4_Susut
Dim cKlik
Dim Titik8(4), Titik7(4), TITIK4(4), titik5(4), Titik6(4), Titik3(4), Titik2(4), Titik(4), Titik1(4)
Dim ccMin(2), ccMax(2), ccTarif(2), ccTKP(2)
Dim PBBMin
Dim ccProses
Private Sub cmdBangunan_Click()
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
frmOP_Tanah.Show
End Sub

Private Sub ccKec_Click()
CALL_KEL
End Sub

Private Sub ccKec_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

If InStr("0123456789-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
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

If InStr("0123456789-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
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

Private Sub ccTahun_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
        If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
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

Private Sub cmdCear_Click()
On Error Resume Next
ccTahun.Text = ccTahun.List(0)
ccKec.Text = ""
ccKel.Text = ""
tJum.Text = 0
xPro(0).Text = 0
xPro(1).Text = ""
LNilai.Visible = False
vBangunan.ListItems.Clear
vBng.ListItems.Clear
vFAS.ListItems.Clear
vOP.ListItems.Clear
ccProses = 0
pNilai.Visible = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
vBangunan.ListItems.Clear
vBng.ListItems.Clear
vFAS.ListItems.Clear
vOP.ListItems.Clear
ccProses = 0
If ccKec.Text = "" And ccKel.Text = "" Then
    xxTanya = MsgBox("Kecamatan Belum Dipilih..." & _
            vbCrLf & " Proses Keseluruhan Sekaligus?", vbCritical + vbYesNo, "Tetnong...")
    If xxTanya = vbNo Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ccProses = 1
End If
If ccKec.Text <> "" And ccKel.Text = "" Then
    ttanya = MsgBox("Anda Belum Memilih Kelurahan, " & _
            vbCrLf & "Proses Seluruh Kelurahan?", vbCritical + vbYesNo, "Tetnong...")
    If ttanya = vbNo Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ccProses = 2
End If

'Timer1.Enabled = True
If ccProses = 1 Then
    T_SQL = "select * from SPPT WHERE THN_PAJAK_SPPT='" & ccTahun.Text & "'"
ElseIf ccProses = 2 Then
    T_SQL = "select * from SPPT WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'"
Else
    T_SQL = "select * from SPPT WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'"
End If
openDB (T_SQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then
    xTanya = MsgBox("Wilayah yang anda pilih sudah dinilai, " & _
            vbCrLf & " Lakukan penilaian ulang??", vbCritical + vbYesNo, "Tetnong...")
    If xTanya = vbNo Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'Hapus SPPT Yang sudah dinilai
    If ccProses = 1 Then
        T_SQL = "DELETE FROM SPPT WHERE THN_PAJAK_SPPT='" & ccTahun.Text & "'"
    ElseIf ccProses = 2 Then
        T_SQL = "DELETE FROM SPPT WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'"
    Else
        T_SQL = "DELETE FROM SPPT WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'"
    End If
    openDB (T_SQL)
    'Hapus SPPT Yang sudah dibayar
    If ccProses = 1 Then
        D_SQL = "DELETE FROM PEMBAYARAN_SPPT WHERE THN_PAJAK_SPPT='" & ccTahun.Text & "'"
    ElseIf ccProses = 2 Then
        D_SQL = "DELETE FROM PEMBAYARAN_SPPT WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'"
    Else
        D_SQL = "DELETE FROM PEMBAYARAN_SPPT WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'"
    End If
    openDB (D_SQL)
End If

TANYA = MsgBox("Proses ini tidak boleh di CANCEL/REJEK!," & _
        vbCrLf & "Apakah dilanjutkan? ", vbInformation + vbYesNo, " Confirmed...")
If TANYA = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
NILAI_BUMI
tJum.Refresh
tJum.Text = vBangunan.ListItems.Count
NILAI_BANGUNAN
NILAI_FASILITAS
DBKB_FAS1A
DBKB_FAS3A
DBKB_FAS2A
NFAS
NILAI_INDIVIDU
TIDAK_KENA_PAJAK
'TAMPIL_OP

'Dim xTarif As Single
'callNIR
'callTarif
'hitKelas1 (xNIR)
'tBumi(2).Text = xKelas
'tBumi(3).Text = Format(tBumi(1).Text * xNilai_Kelas, "#,#0")
'BatasTarif (tBumi(3).Text)
'
'yTarif = xxTarif * tBumi(3).Text
'        If yTarif < PBBMin Then
'            tBumi(4).Text = PBBMin
'        Else
'            tBumi(4).Text = Format(yTarif, "#,#0.00")
'        End If
'    LTarif.Caption = xxTarif
'LKelas.Caption = "NIR: " & xNIR 'tampil di luar form/lebarin az form
'LNilai.Caption = "Nilai Kelas :" & xNilai_Kelas
'----Cek Database SPPT Apakah Sudah Dimasukkan Atau Belum
mnCetak1_Click
vOP.ListItems.Clear
CALL_OP

'Simpan Data Bumi, Bangunan, Data Individu dan Data Objek Pajak
sv_bumi
sv_bangunan
sv_individu
sv_Objek
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

'Private Sub Command1_Click()
'On Error Resume Next
'Dim eAPP As Excel.Application
'Dim eBuku As Excel.Workbook
'Dim eWS As Excel.Worksheet
'Dim brs, datKe As Integer
'Set eAPP = New Excel.Application
'Set eBuku = eAPP.Workbooks.Add
'With eAPP
'    .StandardFontSize = "10"
'End With
'eAPP.Visible = True
'Set eWS = eBuku.Worksheets(1)
'eWS.Select
'Range("A1:B1").Select
'Selection.MergeCells = True
'Selection.HorizontalAlignment = xlCenter
'ActiveCell.FormulaR1C1 = UCase("Data PBB")
'Selection.Font.Bold = True
'Selection.Font.Name = "Verdana"
''With eWS
''    .Cells(2, 1).Value = "Kode"
''    .Cells(2, 2).Value = "Nama"
''    Label1.Caption = "Status : Processing Data..."
''    brs = 100
''    datKe = 0
''    If Not dataPBB.EOF Then
''        dataPBB.MoveFirst
''        While Not dataPBB.EOF
''            Label1.Caption = "Status Exporting data ke " & datKe
''            Label1.Refresh
''            datKe = datKe + 1
''            .Cells(1, 5).Value = "Fetching data ke " & datKe
''            .Cells(brs, 1) = dataPBB!KD_ADJ
''            .Cells(brs, 2) = dataPBB!NM_ADJ
''            brs = brs + 1
''            dataPBB.MoveNext
''        Wend
''    End If
''    .Cells(1, 5).ClearContents
''    .Columns("A:A").EntireColumn.AutoFit
''    .Columns("B:B").EntireColumn.AutoFit
''End With
''dataPBB.Close
''Label1.Caption = "Status : Selesai"
'''On Error GoTo 0
''Set eWS = Nothing
''Set eBuku = Nothing
''eAPP.Quit
'With eWS
'    .Cells(2, 1).Value = "No"
'    .Cells(2, 2).Value = "JPB"
'    .Cells(2, 3).Value = "NOP"
'    .Cells(2, 4).Value = "Nilai Sistem"
'    .Cells(2, 5).Value = "Nilai Manual"
'    Label1.Caption = "Status : Processing Data..."
'    brs = 3
'    datKe = 0
'    LK = 500
'        Do While brs < LK
'            Label1.Caption = "Status Exporting data ke " & datKe
'            Label1.Refresh
'            datKe = datKe + 1
'            .Cells(1, 5).Value = "Fetching data ke " & datKe
'            .Cells(brs, 1) = vBng.ListItems.Item(datKe).ListSubItems(1).Text 'DataGrid1.Columns(0).Text
'            .Cells(brs, 2) = vBng.ListItems.Item(datKe).ListSubItems(10).Text 'DataGrid1.Columns(2).Text
'            .Cells(brs, 3) = vBng.ListItems.Item(datKe).ListSubItems(64).Text 'DataGrid1.Columns(2).Text
'            .Cells(brs, 4) = vBng.ListItems.Item(datKe).ListSubItems(21).Text 'DataGrid1.Columns(2).Text
'            .Cells(brs, 5) = vBng.ListItems.Item(datKe).ListSubItems(63).Text 'DataGrid1.Columns(2).Text
'            brs = brs + 1
'        Loop
'
'    .Cells(1, 5).ClearContents
'    .Columns("A:A").EntireColumn.AutoFit
'    .Columns("B:B").EntireColumn.AutoFit
'End With
'
'Label1.Caption = "Status : Selesai"
''On Error GoTo 0
'
'eAPP.Quit
'End Sub

Private Sub cPro_Click()
On Error Resume Next
If cPro.Value = 1 Then
    Me.Width = 13020: Me.Height = 3960
Else
    Me.Width = 6120: Me.Height = 3960
End If
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
End Sub

Private Sub Form_Activate()
On Error Resume Next
cPro.Value = 0
Screen.MousePointer = vbHourglass
Me.Width = 6120: Me.Height = 3960
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
QSTR1 = "SELECT THN_AWAL_KLS_TANAH FROM KELAS_TANAH order by THN_AWAL_KLS_TANAH desc"
openDB (QSTR1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
xTT = rPajak!THN_AWAL_KLS_TANAH
QSTR2 = "SELECT THN_AWAL_KLS_BNG FROM KELAS_BANGUNAN order by THN_AWAL_KLS_BNG desc"
openDB (QSTR2)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
xTB = rPajak!THN_AWAL_KLS_BNG

ccTahun.Clear
ccTahun.Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    ccTahun.AddItem i
Next
CALL_KEC
'vOP.Visible = False
Timer1.Enabled = False
LNilai.Visible = False
pNilai.Visible = False
Screen.MousePointer = vbDefault
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
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
Screen.MousePointer = vbDefault
End Sub

'Sub callNIR()
'strKab = "Select * From DAT_NIR where KD_ZNT = '" & tBumi(6).Text & "' and THN_NIR_ZNT='" & Trim(cboNOP(5).Text) - 1 & "' and KD_KECAMATAN='" & Left(cboNOP(0).Text, 3) & "' and KD_KELURAHAN='" & Left(cboNOP(1).Text, 3) & "' order by THN_NIR_ZNT,KD_ZNT asc"
'openDB (strKab)
'If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'Do While Not rPajak.EOF
'    tBumi(2).Text = Format(Trim(rPajak!NIR) * 1000, "#,#0")
'        xNIR = Format(Trim(rPajak!NIR) * 1000, "#,#0")
'rPajak.MoveNext
'Loop
'End Sub
Sub callTarif()
On Error GoTo Salah
strTarif = "Select * From TARIF order by NJOP_MIN Asc"
openDB (strTarif)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
Do While Not rPajak.EOF
    i = i + 1
    ccMin(i) = rPajak!NJOP_MIN
    ccMax(i) = rPajak!NJOP_MAX
    ccTarif(i) = rPajak!NILAI_TARIF
    ccTKP(i) = rPajak!NJOPTKP
rPajak.MoveNext
Loop
strTarif = "Select * From PBB_MINIMAL Where THN_PBB_MINIMAL>='" & ccTahun.Text & "'"
openDB (strTarif)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then
    PBBMin = rPajak!NILAI_PBB_MINIMAL
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
xxID = 0
End Sub
Sub hitKelas1(BUMI As Single)
On Error GoTo Salah
    StrQ = "Select * From KELAS_TANAH WHERE THN_AWAL_KLS_TANAH>='" & xTT & "'"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        If BUMI * 0.001 >= rPajak!NILAI_MIN_TANAH * 1 And BUMI * 0.001 <= rPajak!NILAI_MAX_TANAH * 1 Then
            xKelas = rPajak!KD_KLS_TANAH
            xNilai_Kelas = rPajak!NILAI_PER_M2_TANAH * 1000
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
Sub NILAI_BUMI()
On Error GoTo Salah
Timer1.Enabled = True

vBangunan.ListItems.Clear
    If ccProses = 1 Then
        StrQ = "SELECT DAT_OP_BUMI.KD_PROPINSI, DAT_OP_BUMI.KD_DATI2, DAT_OP_BUMI.KD_KECAMATAN, DAT_OP_BUMI.KD_KELURAHAN, DAT_OP_BUMI.KD_BLOK, DAT_OP_BUMI.NO_URUT, DAT_OP_BUMI.KD_JNS_OP, DAT_OP_BUMI.NO_BUMI, DAT_OP_BUMI.KD_ZNT, DAT_NIR.NIR, DAT_OP_BUMI.LUAS_BUMI, DAT_OP_BUMI.JNS_BUMI, DAT_OP_BUMI.NILAI_SISTEM_BUMI, DAT_NIR.[THN_NIR_ZNT]" & _
            "FROM DAT_NIR INNER JOIN DAT_OP_BUMI ON (DAT_NIR.KD_ZNT = DAT_OP_BUMI.KD_ZNT) AND (DAT_NIR.KD_KELURAHAN = DAT_OP_BUMI.KD_KELURAHAN) AND (DAT_NIR.KD_KECAMATAN = DAT_OP_BUMI.KD_KECAMATAN) WHERE DAT_NIR.[THN_NIR_ZNT]='" & ccTahun.Text & "' "
    ElseIf ccProses = 2 Then
        StrQ = "SELECT DAT_OP_BUMI.KD_PROPINSI, DAT_OP_BUMI.KD_DATI2, DAT_OP_BUMI.KD_KECAMATAN, DAT_OP_BUMI.KD_KELURAHAN, DAT_OP_BUMI.KD_BLOK, DAT_OP_BUMI.NO_URUT, DAT_OP_BUMI.KD_JNS_OP, DAT_OP_BUMI.NO_BUMI, DAT_OP_BUMI.KD_ZNT, DAT_NIR.NIR, DAT_OP_BUMI.LUAS_BUMI, DAT_OP_BUMI.JNS_BUMI, DAT_OP_BUMI.NILAI_SISTEM_BUMI, DAT_NIR.[THN_NIR_ZNT]" & _
                "FROM DAT_NIR INNER JOIN DAT_OP_BUMI ON (DAT_NIR.KD_ZNT = DAT_OP_BUMI.KD_ZNT) AND (DAT_NIR.KD_KELURAHAN = DAT_OP_BUMI.KD_KELURAHAN) AND (DAT_NIR.KD_KECAMATAN = DAT_OP_BUMI.KD_KECAMATAN) WHERE DAT_NIR.[THN_NIR_ZNT]='" & ccTahun.Text & "' AND DAT_OP_BUMI.KD_KECAMATAN='" & Left(ccKec.Text, 3) & "'"
    Else
        StrQ = "SELECT DAT_OP_BUMI.KD_PROPINSI, DAT_OP_BUMI.KD_DATI2, DAT_OP_BUMI.KD_KECAMATAN, DAT_OP_BUMI.KD_KELURAHAN, DAT_OP_BUMI.KD_BLOK, DAT_OP_BUMI.NO_URUT, DAT_OP_BUMI.KD_JNS_OP, DAT_OP_BUMI.NO_BUMI, DAT_OP_BUMI.KD_ZNT, DAT_NIR.NIR, DAT_OP_BUMI.LUAS_BUMI, DAT_OP_BUMI.JNS_BUMI, DAT_OP_BUMI.NILAI_SISTEM_BUMI, DAT_NIR.[THN_NIR_ZNT]" & _
                "FROM DAT_NIR INNER JOIN DAT_OP_BUMI ON (DAT_NIR.KD_ZNT = DAT_OP_BUMI.KD_ZNT) AND (DAT_NIR.KD_KELURAHAN = DAT_OP_BUMI.KD_KELURAHAN) AND (DAT_NIR.KD_KECAMATAN = DAT_OP_BUMI.KD_KECAMATAN) WHERE DAT_NIR.[THN_NIR_ZNT]='" & ccTahun.Text & "' AND DAT_OP_BUMI.KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND DAT_OP_BUMI.KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'"
    End If

    'STRQ = "NILAI_BUMI '" & ccTahun.Text & "','" & ccProses & "', '" & Left(ccKec.Text, 3) & "','" & Left(ccKel.Text, 3) & "'"
    openDB (StrQ)
'    frmBar.Show
'    frmBar.Bar1.Max = rPajak.RecordCount + 1
'    frmBar.Bar1.Min = 1
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    tJum.Refresh
        tJum.Text = rPajak.RecordCount
        If rPajak.RecordCount > 1 Then
            pNilai.Max = rPajak.RecordCount
            pNilai.Min = 1
        End If
        pNilai.Visible = True
          i = 0: datKe = 0: J = 0
        Do While Not rPajak.EOF
        i = i + 1
        
        
'            xPro(1).Text = datKe
'            xPro(1).Refresh
'            datKe = datKe + 1
        LNilai.Visible = True
        LNilai.Caption = "1/13 - Proses Penilaian Bumi: " & Round(i / pNilai.Max * 100, 0) & "%" '" & Titik(J)
        LNilai.Refresh
        pNilai.Value = i
        LNilai.Visible = False
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
        xPro(0).Text = i
        xPro(0).Refresh
        xPro(1).Text = xNOP
        xPro(1).Refresh
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_PROPINSI])
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![KD_DATI2])
        vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![KD_KECAMATAN])
        vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Trim(rPajak![KD_KELURAHAN])
        vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", Trim(rPajak![KD_BLOK])
         vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", Trim(rPajak![NO_URUT])
        vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", Trim(rPajak![KD_JNS_OP])
        vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", Trim(rPajak![NO_BUMI])
        vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", Trim(rPajak![KD_ZNT])
        vBangunan.ListItems.Item(i).ListSubItems.Add 11, "", Trim(rPajak![NIR])
        vBangunan.ListItems.Item(i).ListSubItems.Add 12, "", Trim(rPajak![LUAS_BUMI])
        vBangunan.ListItems.Item(i).ListSubItems.Add 13, "", Trim(rPajak![JNS_BUMI])
        vBangunan.ListItems.Item(i).ListSubItems.Add 14, "", Trim(rPajak![NIR]) * Trim(rPajak![LUAS_BUMI]) 'Trim(rPajak![NILAI_SISTEM_BUMI])
        vBangunan.ListItems.Item(i).ListSubItems.Add 15, "", 0 'Trim(rPajak![JNS_BUMI])
        vBangunan.ListItems.Item(i).ListSubItems.Add 16, "", 0 'Trim(rPajak![JNS_BUMI])
        vBangunan.ListItems.Item(i).ListSubItems.Add 17, "", xNOP
                
                'frmBar.Bar1.Value = frmBar.Bar1.Value + 1
    rPajak.MoveNext
    Loop
    judul_nilai_bumi
        CALL_KELAS
        'Set DataGrid1.DataSource = rPajak
    Timer1.Enabled = False
    
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
Sub CALL_KELAS()
On Error GoTo Salah
QSTR = "SELECT * FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH='" & xTT & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    For J = 1 To vBangunan.ListItems.Count
    If vBangunan.ListItems.Item(J).ListSubItems(11).Text * 1 >= rPajak!NILAI_MIN_TANAH And vBangunan.ListItems.Item(J).ListSubItems(11).Text * 1 <= rPajak!NILAI_MAX_TANAH Then
        vBangunan.ListItems.Item(J).ListSubItems(15).Text = Format(rPajak!KD_KLS_TANAH, "000")
        vBangunan.ListItems.Item(J).ListSubItems(16).Text = Format(rPajak!NILAI_PER_M2_TANAH * vBangunan.ListItems.Item(J).ListSubItems(12).Text * 1000, "#,#0.00")
    End If
    Next
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Sub NILAI_BANGUNAN()
On Error GoTo Salah

Dim xTahun
Timer1.Enabled = True

vBng.ListItems.Clear
    'StrQ = "SELECT DAT_OP_BUMI.KD_PROPINSI, DAT_OP_BUMI.KD_DATI2, DAT_OP_BUMI.KD_KECAMATAN, DAT_OP_BUMI.KD_KELURAHAN, DAT_OP_BUMI.KD_BLOK, DAT_OP_BUMI.NO_URUT, DAT_OP_BUMI.KD_JNS_OP, DAT_OP_BUMI.NO_BUMI, DAT_OP_BUMI.KD_ZNT, DAT_NIR.NIR, DAT_OP_BUMI.LUAS_BUMI, DAT_OP_BUMI.JNS_BUMI, DAT_OP_BUMI.NILAI_SISTEM_BUMI, DAT_NIR.[THN_NIR_ZNT]" & _
            "FROM DAT_NIR INNER JOIN DAT_OP_BUMI ON (DAT_NIR.KD_ZNT = DAT_OP_BUMI.KD_ZNT) AND (DAT_NIR.KD_KELURAHAN = DAT_OP_BUMI.KD_KELURAHAN) AND (DAT_NIR.KD_KECAMATAN = DAT_OP_BUMI.KD_KECAMATAN) WHERE DAT_NIR.[THN_NIR_ZNT]='" & ccTahun.Text & "' AND DAT_OP_BUMI.KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND DAT_OP_BUMI.KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'"
    If ccProses = 1 Then
        StrQ = "SELECT * FROM DAT_OP_BANGUNAN "
    ElseIf ccProses = 2 Then
        StrQ = "SELECT * FROM DAT_OP_BANGUNAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' "
    Else
        StrQ = "SELECT * FROM DAT_OP_BANGUNAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'"
    End If
    
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    frmBar.Show
'    frmBar.Bar1.Max = rPajak.RecordCount
'    frmBar.Bar1.Min = 1
        If rPajak.RecordCount > 1 Then
            pNilai.Max = rPajak.RecordCount
            pNilai.Min = 1
        End If
        tJum.Refresh
        tJum.Text = rPajak.RecordCount
          i = 0
    Do While Not rPajak.EOF
        i = i + 1
        LNilai.Visible = True
        LNilai.Caption = "2/13 - Proses Penilaian Bangunan: " & Round(i / pNilai.Max * 100, 0) & "%" '" & Titik1(J)
        pNilai.Value = i
        LNilai.Refresh
        LNilai.Visible = False
        vBng.ListItems.Add i, "", Format(i, "#")
        vBng.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#,#0")
        xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
        xPro(0).Text = i
        xPro(0).Refresh
        xPro(1).Text = xNOP
        xPro(1).Refresh
        vBng.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_PROPINSI])
        vBng.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![KD_DATI2])
        vBng.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![KD_KECAMATAN])
        vBng.ListItems.Item(i).ListSubItems.Add 5, "", Trim(rPajak![KD_KELURAHAN])
        vBng.ListItems.Item(i).ListSubItems.Add 6, "", Trim(rPajak![KD_BLOK])
          vBng.ListItems.Item(i).ListSubItems.Add 7, "", Trim(rPajak![NO_URUT])
        vBng.ListItems.Item(i).ListSubItems.Add 8, "", Trim(rPajak![KD_JNS_OP])
        vBng.ListItems.Item(i).ListSubItems.Add 9, "", Trim(rPajak![NO_BNG])
        vBng.ListItems.Item(i).ListSubItems.Add 10, "", Trim(rPajak![KD_JPB])
        vBng.ListItems.Item(i).ListSubItems.Add 11, "", Trim(rPajak![THN_DIBANGUN_BNG])
        '------------
        xTahun = Trim(rPajak![THN_RENOVASI_BNG])
        If IsNull(xTahun) = True Or xTahun = "" Or xTahun = "-" Then
            xTahun = 0
        End If
        vBng.ListItems.Item(i).ListSubItems.Add 12, "", xTahun
        If IsNull(rPajak![LUAS_BNG]) = True Or Trim(rPajak![LUAS_BNG]) = "" Then
            vBng.ListItems.Item(i).ListSubItems.Add 13, "", 0
        Else
            vBng.ListItems.Item(i).ListSubItems.Add 13, "", Trim(rPajak![LUAS_BNG])
        End If
        If IsNull(rPajak![JML_LANTAI_BNG]) = True Or Trim(rPajak![JML_LANTAI_BNG]) = "" Then
            vBng.ListItems.Item(i).ListSubItems.Add 14, "", 0
            XLT = 1
        Else
            vBng.ListItems.Item(i).ListSubItems.Add 14, "", Trim(rPajak![JML_LANTAI_BNG])
            XLT = Trim(rPajak![JML_LANTAI_BNG])
        End If
        If IsNull(rPajak![KONDISI_BNG]) = True Or Trim(rPajak![KONDISI_BNG]) = "" Then
            vBng.ListItems.Item(i).ListSubItems.Add 15, "", 0
        Else
            vBng.ListItems.Item(i).ListSubItems.Add 15, "", Trim(rPajak![KONDISI_BNG])
        End If
        If IsNull(rPajak![JNS_KONSTRUKSI_BNG]) = True Or Trim(rPajak![JNS_KONSTRUKSI_BNG]) = "" Then
            vBng.ListItems.Item(i).ListSubItems.Add 16, "", 0
        Else
            vBng.ListItems.Item(i).ListSubItems.Add 16, "", Trim(rPajak![JNS_KONSTRUKSI_BNG])
        End If
        If IsNull(rPajak![JNS_ATAP_BNG]) = True Or Trim(rPajak![JNS_ATAP_BNG]) = "" Then
            vBng.ListItems.Item(i).ListSubItems.Add 17, "", 0
        Else
            vBng.ListItems.Item(i).ListSubItems.Add 17, "", Trim(rPajak![JNS_ATAP_BNG])
        End If
        If IsNull(rPajak![KD_DINDING]) = True Or Trim(rPajak![KD_DINDING]) = "" Then
            vBng.ListItems.Item(i).ListSubItems.Add 18, "", 0
        Else
            vBng.ListItems.Item(i).ListSubItems.Add 18, "", Trim(rPajak![KD_DINDING])
        End If
        If IsNull(rPajak![KD_LANTAI]) = True Or Trim(rPajak![KD_LANTAI]) = "" Then
            vBng.ListItems.Item(i).ListSubItems.Add 19, "", 0
        Else
            vBng.ListItems.Item(i).ListSubItems.Add 19, "", Trim(rPajak![KD_LANTAI])
        End If
        If IsNull(rPajak![KD_LANGIT_LANGIT]) = True Or Trim(rPajak![KD_LANGIT_LANGIT]) = "" Then
            vBng.ListItems.Item(i).ListSubItems.Add 20, "", 0
        Else
            vBng.ListItems.Item(i).ListSubItems.Add 20, "", Trim(rPajak![KD_LANGIT_LANGIT])
        End If
        vBng.ListItems.Item(i).ListSubItems.Add 21, "", Format(rPajak![NILAI_SISTEM_BNG], "#,#0.00")
        vBng.ListItems.Item(i).ListSubItems.Add 22, "", 0 'Trim(rPajak![JNS_BUMI])
        vBng.ListItems.Item(i).ListSubItems.Add 23, "", 0 'Trim(rPajak![JNS_BUMI])
        vBng.ListItems.Item(i).ListSubItems.Add 24, "", 0 'Format(rPajak![NILAI_SISTEM_BNG] / rPajak![LUAS_BNG], "#,#0.00")
        vBng.ListItems.Item(i).ListSubItems.Add 25, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 26, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 27, "", 0
        '=======
        vBng.ListItems.Item(i).ListSubItems.Add 28, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 29, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 30, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 31, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 32, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 33, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 34, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 35, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 36, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 37, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 38, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 39, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 40, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 41, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 42, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 43, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 44, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 45, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 46, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 47, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 48, "", 0
        '===========
        vBng.ListItems.Item(i).ListSubItems.Add 49, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 50, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 51, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 52, "", 0
        '-------------
        vBng.ListItems.Item(i).ListSubItems.Add 53, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 54, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 55, "", 0
        '==================
        vBng.ListItems.Item(i).ListSubItems.Add 56, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 57, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 58, "", 0
        vBng.ListItems.Item(i).ListSubItems.Add 59, "", 0
        '==========
        vBng.ListItems.Item(i).ListSubItems.Add 60, "", 0 'Nilai Fasilitas, tidak mengalami susut
        '==========
        vBng.ListItems.Item(i).ListSubItems.Add 61, "", 0 'Nilai Fasilitas, mengalami susut
        '================
        vBng.ListItems.Item(i).ListSubItems.Add 62, "", 0 'Nilai Total Sistem
        '==========
    
        
        vBng.ListItems.Item(i).ListSubItems.Add 63, "", 0 'Nilai Susut
        vBng.ListItems.Item(i).ListSubItems.Add 64, "", 0 'NOP
        
        vBng.ListItems.Item(i).ListSubItems.Add 65, "", 0 'Nilai Susut
        
        vBng.ListItems.Item(i).ListSubItems.Add 66, "", 0 'NILAI RUANGAN KAMAR
        vBng.ListItems.Item(i).ListSubItems.Add 67, "", 0 'NILAI RUANGAN LAIN
        vBng.ListItems.Item(i).ListSubItems.Add 68, "", 0 'BOILER
        vBng.ListItems.Item(i).ListSubItems.Add 69, "", "SISTEM" 'KETERANGAN APAKAH NILAI SISTEM ATAU INDIVIDU
        vBng.ListItems.Item(i).ListSubItems.Add 70, "", 0 'DBKB Fasilitas Disusutkan
        vBng.ListItems.Item(i).ListSubItems.Add 71, "", 0 'Nilai Keliling Dinding Daya Dukung Lantai JPB 8 dan JPB 3
        
        vBng.ListItems.Item(i).ListSubItems.Add 72, "", rPajak!NO_FORMULIR_LSPOP 'Nomor Formulir
        vBng.ListItems.Item(i).ListSubItems.Add 73, "", rPajak!JNS_TRANSAKSI_BNG 'Jenis Transaksi
        If IsNull(rPajak!NIP_PENDATA_BNG) = True Or rPajak!NIP_PENDATA_BNG = "" Then rPajak!NIP_PENDATA_BNG = "-"
        If IsNull(rPajak!NIP_PEMERIKSA_BNG) = True Or rPajak!NIP_PEMERIKSA_BNG = "" Then rPajak!NIP_PEMERIKSA_BNG = "-"
        If IsNull(rPajak!NIP_PEREKAM_BNG) = True Or rPajak!NIP_PEREKAM_BNG = "" Then rPajak!NIP_PEREKAM_BNG = "-"


        vBng.ListItems.Item(i).ListSubItems.Add 74, "", rPajak![TGL_PENDATAAN_BNG] 'Tanggal Pendataan
        vBng.ListItems.Item(i).ListSubItems.Add 75, "", rPajak![NIP_PENDATA_BNG] 'NIP Petugas Pendata
        vBng.ListItems.Item(i).ListSubItems.Add 76, "", rPajak![TGL_PEMERIKSAAN_BNG] 'Tanggal Pemeriksaan
        vBng.ListItems.Item(i).ListSubItems.Add 77, "", rPajak![NIP_PEMERIKSA_BNG] 'NIP Petugas Pemeriksa
        vBng.ListItems.Item(i).ListSubItems.Add 78, "", rPajak![TGL_PEREKAMAN_BNG]  'Tanggal Perekaman
        vBng.ListItems.Item(i).ListSubItems.Add 79, "", rPajak![NIP_PEREKAM_BNG] 'NIP Petugas Perekam
        vBng.ListItems.Item(i).ListSubItems.Add 80, "", 0 'NILAI INDIVIDUAL
        vBng.ListItems.Item(i).ListSubItems.Add 81, "", 0 'Biaya Komponen Utama
        vBng.ListItems.Item(i).ListSubItems.Add 82, "", 0 'Biaya Komponen Material
        vBng.ListItems.Item(i).ListSubItems.Add 83, "", 0 'Biaya Fasilitas Yang Disusutkan
        vBng.ListItems.Item(i).ListSubItems.Add 84, "", 0 'Jumlah Persentase Susut
        vBng.ListItems.Item(i).ListSubItems.Add 85, "", 0 'Nilai Penyusutan
        vBng.ListItems.Item(i).ListSubItems.Add 86, "", 0 'Biaya Fasilitas Yang Tidak Disusutkan
        
              
                'frmBar.Bar1.Value = i
    rPajak.MoveNext
    Loop
    judul_nilai_bng
        'CALL_KELAS1
        'CALL_NJOPTKP
        CALL_DBKB_STANDARD
        DAYA_DUKUNG_JPB3
        CALL_DAYA_DKUNG_JPB3
        DAYA_DUKUNG_JPB8
        CALL_DBKB_JPB3
        CALL_DBKB_JPB2
        CALL_DBKB_JPB4
        CALL_DBKB_JPB5
        CALL_DBKB_JPB6
        CALL_DBKB_JPB7
        CALL_DBKB_JPB8
        CALL_DBKB_JPB9
        CALL_DBKB_JPB12
        CALL_DBKB_JPB13
        CALL_DBKB_JPB14
        CALL_DBKB_JPB15
        CALL_DBKB_JPB16
        CALL_DBKB_JPB17
        MATERIAL
        'NILAI_INDIVIDU
        Timer1.Enabled = False
     
'salah:
'If Err.Number = 0 Then Exit Sub
'MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
Sub CALL_KELAS1()
On Error GoTo Salah
QSTR = "SELECT * FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG ='" & xTB & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    For J = 1 To vBng.ListItems.Count
    If (vBng.ListItems.Item(J).ListSubItems(21).Text * 1 / vBng.ListItems.Item(J).ListSubItems(13).Text) >= rPajak!NILAI_MIN_BNG And (vBng.ListItems.Item(J).ListSubItems(21).Text * 1 / vBng.ListItems.Item(J).ListSubItems(13).Text) <= rPajak!NILAI_MAX_BNG Then
        vBng.ListItems.Item(J).ListSubItems(22).Text = Format(rPajak!KD_KLS_BNG, "000")
        vBng.ListItems.Item(J).ListSubItems(23).Text = Format(rPajak!NILAI_PER_M2_BNG, "#,#0.00")
        vBng.ListItems.Item(J).ListSubItems(24).Text = Format(rPajak!NILAI_PER_M2_BNG * vBng.ListItems.Item(J).ListSubItems(13).Text, "#,#0.00")
    End If
    Next
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Sub CALL_NJOPTKP()
On Error GoTo Salah
StrQ = "SELECT * FROM TARIF"
openDB (StrQ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    For J = 1 To vBng.ListItems.Count
    If vBng.ListItems.Item(J).ListSubItems(24).Text * 1000 >= rPajak!NJOP_MIN And vBng.ListItems.Item(J).ListSubItems(24).Text * 1000 <= rPajak!NJOP_MAX Then
        vBng.ListItems.Item(J).ListSubItems(25).Text = Format(rPajak!NJOPTKP, "#,#0.00")
    End If
    Next
    
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub mnCetak1_Click()
On Error GoTo Salah
If cKlik = 0 Then
    vOP.Visible = True
    cKlik = 1
Else
    cKlik = 0
    'vOP.Visible = False
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub Timer1_Timer()
'NILAI_BUMI
'NILAI_BANGUNAN
'Timer = Timer + 1
End Sub

Private Sub vBangunan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan.SortKey = ColumnHeader.Index - 1
vBangunan.Sorted = True
vBangunan.Sorted = False
vBangunan.SortOrder = lvwAscending
End Sub

Private Sub vBng_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBng.SortKey = ColumnHeader.Index - 1
vBng.Sorted = True
vBng.Sorted = False
vBng.SortOrder = lvwAscending
End Sub

Private Sub vFAS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vFAS.SortKey = ColumnHeader.Index - 1
vFAS.Sorted = True
vFAS.Sorted = False
vFAS.SortOrder = lvwAscending
End Sub
'=====================Menghitung DBKB===============
Sub CALL_DBKB_STANDARD()
On Error GoTo Salah
Dim JPB, LUAS, JLANTAI, JL
Dim xJPB, XJLANTAI, XBAGI, LMIN, LMAX

'=======BANGUNAN STANDARD UNTUK JPB 01,02,04,07,09 DENGAN LANTAI <=4
    StrQ = "SELECT DBKB_STANDARD.THN_DBKB_STANDARD, DBKB_STANDARD.KD_JPB, TIPE_BANGUNAN.TIPE_BNG, TIPE_BANGUNAN.NM_TIPE_BNG, TIPE_BANGUNAN.LUAS_MIN_TIPE_BNG, TIPE_BANGUNAN.LUAS_MAX_TIPE_BNG, TIPE_BANGUNAN.FAKTOR_PEMBAGI_TIPE_BNG, DBKB_STANDARD.KD_BNG_LANTAI, DBKB_STANDARD.NILAI_DBKB_STANDARD FROM DBKB_STANDARD INNER JOIN TIPE_BANGUNAN ON DBKB_STANDARD.TIPE_BNG = TIPE_BANGUNAN.TIPE_BNG WHERE (((DBKB_STANDARD.THN_DBKB_STANDARD)='" & ccTahun.Text * 1 & "')) order by KD_JPB,KD_BNG_LANTAI ASC" 'AND DBKB_STANDARD.KD_JPB='" & JPB & "'"
    openDB (StrQ)
    i = 0: K = 0

    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        If rPajak.RecordCount > 1 Then
            pNilai.Max = rPajak.RecordCount
            pNilai.Min = 1
        End If
tJum.Refresh
    tJum.Text = rPajak.RecordCount
    Do While Not rPajak.EOF
        K = K + 1
            LNilai.Visible = True
            LNilai.Caption = "3/13 - Proses DBKB Standar : " & Round(K / pNilai.Max * 100, 0) & "%" '" & Titik1(k)
            LNilai.Refresh
            LNilai.Visible = False
            pNilai.Value = K
            
        For J = 1 To vBng.ListItems.Count
            
            JPB = Trim(vBng.ListItems.Item(J).ListSubItems(10).Text)
            LUAS = vBng.ListItems.Item(J).ListSubItems(13).Text * 1
            JLANTAI = vBng.ListItems.Item(J).ListSubItems(14).Text * 1
            xJPB = rPajak!KD_JPB
            XJLANTAI = Mid(Trim(rPajak!KD_BNG_LANTAI), 3, 1)
            XBAGI = rPajak!FAKTOR_PEMBAGI_TIPE_BNG 'Bandingkan dengan luas
            LMIN = rPajak!LUAS_MIN_TIPE_BNG
            LMAX = rPajak!LUAS_MAX_TIPE_BNG
            
            
            Select Case JPB
            Case "01", "02", "04", "05", "07", "09", "10", "11"
            If JLANTAI <= 2 Then
                JL = 1
            ElseIf JLANTAI <= 4 Then
                JL = 2
            Else
                JL = 3
            End If
            
           If (JPB = "01" Or JPB = "10" Or JPB = "11") And JL <= 2 Then JPB = "01"
            If JPB = "05" And JL <= 2 Then JPB = "05"
            If (JPB = "02" Or JPB = "04" Or JPB = "07" Or JPB = "09") And JL <= 2 Then
                JPB = "02"
            End If

            'MsgBox JPB & ":" & LUAS & ":" & JLANTAI
            'If (JPB = "01" Or JPB = "02" Or JPB = "04" Or JPB = "05" Or JPB = "07" Or JPB = "09") And JL <= 2 Then
            If JL <= 2 Then
                'If JPB = XJPB And (LUAS >= LMIN And LUAS <= LMAX) And JL = XJLANTAI And LUAS >= XBAGI Then
                If rPajak!KD_JPB = JPB And (LUAS * 1 >= rPajak!LUAS_MIN_TIPE_BNG And LUAS * 1 <= rPajak!LUAS_MAX_TIPE_BNG) And Mid(Trim(rPajak!KD_BNG_LANTAI), 3, 1) * 1 = JLANTAI Then   'And LUAS * 1 >= rPajak!FAKTOR_PEMBAGI_TIPE_BNG Then
                    nDBKB = rPajak!NILAI_DBKB_STANDARD
                    cTipe = rPajak!TIPE_BNG & "_" & Mid(Trim(rPajak!KD_BNG_LANTAI), 3, 1) * 1 & "-" & JL
                    vBng.ListItems.Item(J).ListSubItems(26).Text = nDBKB
                    vBng.ListItems.Item(J).ListSubItems(27).Text = cTipe
                    vBng.ListItems.Item(J).ListSubItems(28).Text = nDBKB
                    xNOP = vBng.ListItems.Item(J).ListSubItems(64).Text
'                    Select Case JPB
'                    Case "01": vBng.ListItems.Item(j).ListSubItems(28).Text = nDBKB
'                    Case "02": vBng.ListItems.Item(j).ListSubItems(29).Text = nDBKB
'                    Case "04": vBng.ListItems.Item(j).ListSubItems(31).Text = nDBKB
'                    Case "05": vBng.ListItems.Item(j).ListSubItems(32).Text = nDBKB
'                    Case "07": vBng.ListItems.Item(j).ListSubItems(34).Text = nDBKB
'                    Case "09": vBng.ListItems.Item(j).ListSubItems(36).Text = nDBKB
'                    End Select
                End If
            End If
        End Select
        Next
        xPro(0).Text = K
        xPro(0).Refresh
        xPro(1).Text = xNOP
        xPro(1).Refresh
    rPajak.MoveNext
    Loop
        'frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_DBKB_JPB2()
'jpb=10:Thn_Bangun=11:Thn_reno=12:Luas=13:jlantai=14:Kondisi_bng=15:Konstruksi=16:Atap=17:Dinding=18:Lantai=19:Langi2=20
On Error GoTo Salah
CALL_JPB2
 StrQ = "SELECT * FROM DBKB_JPB2 WHERE THN_DBKB_JPB2='" & ccTahun.Text * 1 & "' ORDER BY KLS_DBKB_JPB2 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "02" Then
            If vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_DBKB_JPB2 And (vBng.ListItems.Item(J).ListSubItems(14).Text * 1 >= rPajak!LANTAI_MIN_JPB2 And vBng.ListItems.Item(J).ListSubItems(14).Text * 1 <= rPajak!LANTAI_MAX_JPB2) Then
                If vBng.ListItems.Item(J).ListSubItems(26).Text > 0 Then vBng.ListItems.Item(J).ListSubItems(29).ForeColor = vbRed
                    vBng.ListItems.Item(J).ListSubItems(29).Text = rPajak!NILAI_DBKB_JPB2
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
        'frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub CALL_DBKB_JPB3()
On Error GoTo Salah
 StrQ = "SELECT * FROM DBKB_JPB3 WHERE THN_DBKB_JPB3='" & ccTahun.Text * 1 & "' ORDER BY LBR_BENT_MIN_DBKB_JPB3 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst

        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            'If JPB38(1).Text >= rPajak!LBR_BENT_MIN_DBKB_JPB3 And JPB38(1).Text <= rPajak!LBR_BENT_MAX_DBKB_JPB3 Then
            If vBng.ListItems.Item(J).ListSubItems(10).Text = "03" Then
                If (vBng.ListItems.Item(J).ListSubItems(49).Text * 1 >= rPajak!LBR_BENT_MIN_DBKB_JPB3 * 1 And vBng.ListItems.Item(J).ListSubItems(49).Text * 1 <= rPajak!LBR_BENT_MAX_DBKB_JPB3 * 1) And (vBng.ListItems.Item(J).ListSubItems(50).Text * 1 >= rPajak!TING_KOLOM_MIN_DBKB_JPB3 And vBng.ListItems.Item(J).ListSubItems(50).Text * 1 <= rPajak!TING_KOLOM_MAX_DBKB_JPB3 * 1) Then
                    nDBKB = rPajak!NILAI_DBKB_JPB3
                    vBng.ListItems.Item(J).ListSubItems(30).Text = rPajak!NILAI_DBKB_JPB3 ' * vBng.ListItems.Item(j).ListSubItems(13).Text 'OK
                End If
            End If
        Next
        rPajak.MoveNext
        Loop
        'frmBar.Bar1.Value = frmBar.Bar1.Value
      '  If I = 0 Then nDBKB = 0
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub CALL_DBKB_JPB4()
On Error GoTo Salah
CALL_JPB4
 StrQ = "SELECT * FROM DBKB_JPB4 WHERE THN_DBKB_JPB4='" & ccTahun.Text * 1 & "' ORDER BY KLS_DBKB_JPB4 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "04" Then
            If vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_DBKB_JPB4 And (vBng.ListItems.Item(J).ListSubItems(14).Text * 1 >= rPajak!LANTAI_MIN_JPB4 And vBng.ListItems.Item(J).ListSubItems(14).Text * 1 <= rPajak!LANTAI_MAX_JPB4) Then
                    If vBng.ListItems.Item(J).ListSubItems(26).Text > 0 Then vBng.ListItems.Item(J).ListSubItems(31).ForeColor = vbRed
                        vBng.ListItems.Item(J).ListSubItems(31).Text = rPajak!NILAI_DBKB_JPB4
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
        'frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_DBKB_JPB5()
On Error GoTo Salah
CALL_JPB5
 StrQ = "SELECT * FROM DBKB_JPB5 WHERE THN_DBKB_JPB5='" & ccTahun.Text * 1 & "' ORDER BY KLS_DBKB_JPB5 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "05" Then
            If vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_DBKB_JPB5 And (vBng.ListItems.Item(J).ListSubItems(14).Text * 1 >= rPajak!LANTAI_MIN_JPB5 And vBng.ListItems.Item(J).ListSubItems(14).Text * 1 <= rPajak!LANTAI_MAX_JPB5) Then
                    If vBng.ListItems.Item(J).ListSubItems(26).Text > 0 Then vBng.ListItems.Item(J).ListSubItems(32).ForeColor = vbRed
                        vBng.ListItems.Item(J).ListSubItems(32).Text = rPajak!NILAI_DBKB_JPB5
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
        'frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_DBKB_JPB6()
On Error GoTo Salah
CALL_JPB6
 StrQ = "SELECT * FROM DBKB_JPB6 WHERE THN_DBKB_JPB6='" & ccTahun.Text * 1 & "' ORDER BY KLS_DBKB_JPB6 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "06" Then
            If vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_DBKB_JPB6 Then
                    vBng.ListItems.Item(J).ListSubItems(33).Text = rPajak!NILAI_DBKB_JPB6
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
'                frmBar.Bar1.Value = frmBar.Bar1.Value

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_DBKB_JPB7()
On Error GoTo Salah
CALL_JPB7
 StrQ = "SELECT * FROM DBKB_JPB7 WHERE THN_DBKB_JPB7='" & ccTahun.Text * 1 & "' ORDER BY JNS_DBKB_JPB7 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "07" Then
            'MsgBox vBng.ListItems.Item(j).ListSubItems(51).Text & ":" & rPajak!JNS_DBKB_JPB7
            If (vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!JNS_DBKB_JPB7 And vBng.ListItems.Item(J).ListSubItems(52).Text = rPajak!BINTANG_DBKB_JPB7) Then
                If (vBng.ListItems.Item(J).ListSubItems(14).Text * 1 >= rPajak!LANTAI_MIN_JPB7 And vBng.ListItems.Item(J).ListSubItems(14).Text * 1 <= rPajak!LANTAI_MAX_JPB7) Then
                    If vBng.ListItems.Item(J).ListSubItems(26).Text > 0 Then vBng.ListItems.Item(J).ListSubItems(34).ForeColor = vbRed
                        vBng.ListItems.Item(J).ListSubItems(34).Text = rPajak!NILAI_DBKB_JPB7
                End If
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
'                frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
Sub CALL_DBKB_JPB8()
On Error GoTo Salah
CALL_DAYA_DKUNG_JPB3
 StrQ = "SELECT * FROM DBKB_JPB8 WHERE THN_DBKB_JPB8='" & ccTahun.Text * 1 & "' ORDER BY LBR_BENT_MIN_DBKB_JPB8 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            'If JPB38(1).Text >= rPajak!LBR_BENT_MIN_DBKB_JPB3 And JPB38(1).Text <= rPajak!LBR_BENT_MAX_DBKB_JPB3 Then
            If vBng.ListItems.Item(J).ListSubItems(10).Text = "08" Then
                
                If (vBng.ListItems.Item(J).ListSubItems(49).Text * 1 >= rPajak!LBR_BENT_MIN_DBKB_JPB8 And vBng.ListItems.Item(J).ListSubItems(49).Text * 1 <= rPajak!LBR_BENT_MAX_DBKB_JPB8) And (vBng.ListItems.Item(J).ListSubItems(50).Text * 1 >= rPajak!TING_KOLOM_MIN_DBKB_JPB8 And vBng.ListItems.Item(J).ListSubItems(50).Text * 1 <= rPajak!TING_KOLOM_MAX_DBKB_JPB8) Then
                   vBng.ListItems.Item(J).ListSubItems(35).Text = rPajak!NILAI_DBKB_JPB8
                End If
            End If
        Next
        rPajak.MoveNext
        Loop
      '  If I = 0 Then nDBKB = 0
'              frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
Sub CALL_DBKB_JPB9()
On Error GoTo Salah
CALL_JPB9
 StrQ = "SELECT * FROM DBKB_JPB9 WHERE THN_DBKB_JPB9='" & ccTahun.Text * 1 & "' ORDER BY KLS_DBKB_JPB9 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "09" Then
            If vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_DBKB_JPB9 Then
                If (vBng.ListItems.Item(J).ListSubItems(14).Text * 1 >= rPajak!LANTAI_MIN_JPB9 And vBng.ListItems.Item(J).ListSubItems(14).Text * 1 <= rPajak!LANTAI_MAX_JPB9) Then
                    If vBng.ListItems.Item(J).ListSubItems(26).Text > 0 Then vBng.ListItems.Item(J).ListSubItems(36).ForeColor = vbRed
                        vBng.ListItems.Item(J).ListSubItems(36).Text = rPajak!NILAI_DBKB_JPB9
                End If
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
'                frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
Sub CALL_DBKB_JPB12()
On Error GoTo Salah
CALL_JPB12
 StrQ = "SELECT * FROM DBKB_JPB12 WHERE THN_DBKB_JPB12='" & ccTahun.Text * 1 & "' ORDER BY TYPE_DBKB_JPB12 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "12" Then
            If vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!TYPE_DBKB_JPB12 Then
                vBng.ListItems.Item(J).ListSubItems(39).Text = rPajak!NILAI_DBKB_JPB12
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
'                frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
Sub CALL_DBKB_JPB13()
On Error GoTo Salah
CALL_JPB13
 StrQ = "SELECT * FROM DBKB_JPB13 WHERE THN_DBKB_JPB13='" & ccTahun.Text * 1 & "' ORDER BY KLS_DBKB_JPB13 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "13" Then
            'MsgBox vBng.ListItems.Item(j).ListSubItems(51).Text & ":" & rPajak!JNS_DBKB_JPB13
            If (vBng.ListItems.Item(J).ListSubItems(14).Text * 1 >= rPajak!LANTAI_MIN_JPB13 And vBng.ListItems.Item(J).ListSubItems(14).Text * 1 <= rPajak!LANTAI_MAX_JPB13) Then
                If (vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_DBKB_JPB13) Then ' And vBng.ListItems.Item(j).ListSubItems(52).Text = rPajak!BINTANG_DBKB_JPB13) Then
                    vBng.ListItems.Item(J).ListSubItems(40).Text = rPajak!NILAI_DBKB_JPB13
                End If
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
'        frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_DBKB_JPB14()
'-----Panggil Nilai DBKB JPB14/Kanopi
On Error GoTo Salah
CALL_JPB14
'------
 StrQ = "SELECT * FROM DBKB_JPB14 WHERE THN_DBKB_JPB14='" & ccTahun.Text * 1 & "'"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "14" Then
            vBng.ListItems.Item(J).ListSubItems(41).Text = rPajak!NILAI_DBKB_JPB14 '* vBng.ListItems.Item(j).ListSubItems(45).Text
        End If
        Next
        rPajak.MoveNext
        Loop
'        frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_DBKB_JPB15()
On Error GoTo Salah
'-----Panggil Nilai Tangki Minyak
CALL_JPB15
'------
 StrQ = "SELECT * FROM DBKB_JPB15 WHERE THN_DBKB_JPB15='" & ccTahun.Text * 1 & "'"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "15" Then
            If (vBng.ListItems.Item(J).ListSubItems(46).Text * 1 >= rPajak!KAPASITAS_MIN_DBKB_JPB15 And vBng.ListItems.Item(J).ListSubItems(46).Text * 1 <= rPajak!KAPASITAS_MAX_DBKB_JPB15) And (vBng.ListItems.Item(J).ListSubItems(45).Text * 1 = rPajak!JNS_TANGKI_DBKB_JPB15) Then
                vBng.ListItems.Item(J).ListSubItems(42).Text = rPajak!NILAI_DBKB_JPB15
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
'        frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_DBKB_JPB16()
On Error GoTo Salah
CALL_JPB16
'------
 StrQ = "SELECT * FROM DBKB_JPB16 WHERE THN_DBKB_JPB16='" & ccTahun.Text * 1 & "' ORDER BY KLS_DBKB_JPB16 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "16" Then
            If (vBng.ListItems.Item(J).ListSubItems(14).Text * 1 >= rPajak!LANTAI_MIN_JPB16 And vBng.ListItems.Item(J).ListSubItems(14).Text * 1 <= rPajak!LANTAI_MAX_JPB16) And (vBng.ListItems.Item(J).ListSubItems(51).Text * 1 = rPajak!KLS_DBKB_JPB16) Then
                vBng.ListItems.Item(J).ListSubItems(43).Text = rPajak!NILAI_DBKB_JPB16
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
'        frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_DBKB_JPB17()
On Error GoTo Salah
CALL_JPB17
'------
 StrQ = "SELECT * FROM DBKB_JPB17 WHERE THN_DBKB_JPB17='" & ccTahun.Text * 1 & "' ORDER BY TINGGI_MIN_JPB17,TINGGI_MAX_JPB17 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
        If vBng.ListItems.Item(J).ListSubItems(10).Text = "17" Then
            If (vBng.ListItems.Item(J).ListSubItems(51).Text * 1 >= rPajak!TINGGI_MIN_JPB17 And vBng.ListItems.Item(J).ListSubItems(15).Text * 1 <= rPajak!TINGGI_MAX_JPB17) Then
                vBng.ListItems.Item(J).ListSubItems(44).Text = rPajak!NILAI_DBKB_JPB17 '* vBng.ListItems.Item(j).ListSubItems(13).Text
            End If
        End If
        Next
        rPajak.MoveNext
        Loop
'        frmBar.Bar1.Value = frmBar.Bar1.Value
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub CALL_JPB2()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB2"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text ' & "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_JPB2
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_JPB4()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB4"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) ' & "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_JPB4
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_JPB5()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB5"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) ' & "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text ' & "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(54).Text = rPajak!LUAS_KMR_JPB5_DGN_AC_SENT
                vBng.ListItems.Item(J).ListSubItems(55).Text = rPajak!LUAS_RNG_LAIN_JPB5_DGN_AC_SENT
                vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_JPB5
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_JPB6()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB6"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_JPB6
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_JPB7()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB7"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!JNS_JPB7
                vBng.ListItems.Item(J).ListSubItems(52).Text = rPajak!BINTANG_JPB7
                vBng.ListItems.Item(J).ListSubItems(53).Text = rPajak!JML_KMR_JPB7
                vBng.ListItems.Item(J).ListSubItems(54).Text = rPajak!LUAS_KMR_JPB7_DGN_AC_SENT
                vBng.ListItems.Item(J).ListSubItems(55).Text = rPajak!LUAS_KMR_LAIN_JPB7_DGN_AC_SENT
                
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_JPB9()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB9"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_JPB9
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_JPB12()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB12"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text ' & "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!TYPE_JPB12
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub CALL_JPB13()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB13"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) ' & "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_JPB13
                'vBng.ListItems.Item(j).ListSubItems(52).Text = rPajak!BINTANG_JPB7
                vBng.ListItems.Item(J).ListSubItems(53).Text = rPajak!JML_JPB13
                vBng.ListItems.Item(J).ListSubItems(54).Text = rPajak!LUAS_JPB13_DGN_AC_SENT
                vBng.ListItems.Item(J).ListSubItems(55).Text = rPajak!LUAS_JPB13_LAIN_DGN_AC_SENT
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_JPB14()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB14"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(45).Text = rPajak!LUAS_KANOPI_JPB14
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub CALL_JPB15()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB15"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(45).Text = rPajak!LETAK_TANGKI_JPB15
                vBng.ListItems.Item(J).ListSubItems(46).Text = rPajak!KAPASITAS_TANGKI_JPB15
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub CALL_JPB16()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB16"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!KLS_JPB16
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_JPB17()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_JPB17"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                vBng.ListItems.Item(J).ListSubItems(51).Text = rPajak!TINGGI_BNG_JPB17
            End If
        Next
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
'==============PANGGIL DATABASE MEZANINE==============
Sub call_Mezz()
On Error GoTo Salah
QSTR = "SELECT * FROM DBKB_MEZANIN WHERE THN_DBKB_MEZANIN='" & ccTahun.Text & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    nMezanin = rPajak!NILAI_DBKB_MEZANIN
    rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
'============== Menentukan Khusus Nilai Daya Dukung dan Mezanin
Sub CALL_DAYA_DKUNG_JPB3()
On Error GoTo Salah
QSTR = "SELECT DBKB_DAYA_DUKUNG.KD_PROPINSI, DBKB_DAYA_DUKUNG.KD_DATI2, DBKB_DAYA_DUKUNG.THN_DBKB_DAYA_DUKUNG, DBKB_DAYA_DUKUNG.TYPE_KONSTRUKSI, DAYA_DUKUNG.DAYA_DUKUNG_LANTAI_MIN_DBKB, DAYA_DUKUNG.DAYA_DUKUNG_LANTAI_MAX_DBKB, DBKB_DAYA_DUKUNG.NILAI_DBKB_DAYA_DUKUNG FROM DBKB_DAYA_DUKUNG INNER JOIN DAYA_DUKUNG ON DBKB_DAYA_DUKUNG.TYPE_KONSTRUKSI = DAYA_DUKUNG.TYPE_KONSTRUKSI WHERE DBKB_DAYA_DUKUNG.THN_DBKB_DAYA_DUKUNG='" & ccTahun.Text & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
Do While Not rPajak.EOF
    For J = 1 To vBng.ListItems.Count
    If vBng.ListItems.Item(J).ListSubItems(47).Text * vBng.ListItems.Item(J).ListSubItems(13).Text >= rPajak!DAYA_DUKUNG_LANTAI_MIN_DBKB And vBng.ListItems.Item(J).ListSubItems(47).Text * vBng.ListItems.Item(J).ListSubItems(13).Text <= rPajak!DAYA_DUKUNG_LANTAI_MAX_DBKB Then
        i = i + 1
        nDUKUNG = rPajak!NILAI_DBKB_DAYA_DUKUNG
        vBng.ListItems.Item(J).ListSubItems(46).Text = vBng.ListItems.Item(J).ListSubItems(45).Text * nMezanin 'vBng.ListItems.Item(j).ListSubItems(13).Text
        vBng.ListItems.Item(J).ListSubItems(48).Text = nDUKUNG * vBng.ListItems.Item(J).ListSubItems(13).Text
    End If
Next
    rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

'==============DBKB DAYA DUKUNG LANTAI, KHUSUS UNTUK JPB=3 PABRIK==============
Sub DAYA_DUKUNG_JPB3()
On Error GoTo Salah
Dim XTIPE, xKOLOM, xBENTANG, xMEZZ, xKELILING, xDUKUNG
call_Mezz
'QSTR = "SELECT DBKB_DAYA_DUKUNG.KD_PROPINSI, DBKB_DAYA_DUKUNG.KD_DATI2, DBKB_DAYA_DUKUNG.THN_DBKB_DAYA_DUKUNG, DBKB_DAYA_DUKUNG.TYPE_KONSTRUKSI, DAYA_DUKUNG.DAYA_DUKUNG_LANTAI_MIN_DBKB, DAYA_DUKUNG.DAYA_DUKUNG_LANTAI_MAX_DBKB, DBKB_DAYA_DUKUNG.NILAI_DBKB_DAYA_DUKUNG FROM DBKB_DAYA_DUKUNG INNER JOIN DAYA_DUKUNG ON DBKB_DAYA_DUKUNG.TYPE_KONSTRUKSI = DAYA_DUKUNG.TYPE_KONSTRUKSI WHERE DBKB_DAYA_DUKUNG.THN_DBKB_DAYA_DUKUNG='" & ccTahun.Text & "'"
QSTR = "Select * FROM DAT_JPB3"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                XTIPE = rPajak!TYPE_KONSTRUKSI
                xKOLOM = rPajak!TING_KOLOM_JPB3
                xBENTANG = rPajak!LBR_BENT_JPB3
                xMEZZ = rPajak!LUAS_MEZZANINE_JPB3
                xKELILING = rPajak!KELILING_DINDING_JPB3
                xDUKUNG = rPajak!DAYA_DUKUNG_LANTAI_JPB3
                'If xDUKUNG * 1 >= rPajak!DAYA_DUKUNG_LANTAI_MIN_DBKB And xDUKUNG * 1 <= rPajak!DAYA_DUKUNG_LANTAI_MAX_DBKB Then
                     vBng.ListItems.Item(J).ListSubItems(45).Text = xMEZZ
                     vBng.ListItems.Item(J).ListSubItems(47).Text = xDUKUNG 'rPajak!NILAI_DBKB_DAYA_DUKUNG * vBng.ListItems.Item(j).ListSubItems(13).Text
                     vBng.ListItems.Item(J).ListSubItems(49).Text = xBENTANG
                     vBng.ListItems.Item(J).ListSubItems(50).Text = xKOLOM
                     vBng.ListItems.Item(J).ListSubItems(71).Text = xKELILING
                'End If
            End If
        Next

    
    rPajak.MoveNext
Loop
If i = 0 Then nDUKUNG = 0
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
'==============DAYA DUKUNG LANTAI, KHUSUS UNTUK JPB=8==============
Sub DAYA_DUKUNG_JPB8()
On Error GoTo Salah
Dim XTIPE, xKOLOM, xBENTANG, xMEZZ, xKELILING, xDUKUNG
call_Mezz
'QSTR = "SELECT DBKB_DAYA_DUKUNG.KD_PROPINSI, DBKB_DAYA_DUKUNG.KD_DATI2, DBKB_DAYA_DUKUNG.THN_DBKB_DAYA_DUKUNG, DBKB_DAYA_DUKUNG.TYPE_KONSTRUKSI, DAYA_DUKUNG.DAYA_DUKUNG_LANTAI_MIN_DBKB, DAYA_DUKUNG.DAYA_DUKUNG_LANTAI_MAX_DBKB, DBKB_DAYA_DUKUNG.NILAI_DBKB_DAYA_DUKUNG FROM DBKB_DAYA_DUKUNG INNER JOIN DAYA_DUKUNG ON DBKB_DAYA_DUKUNG.TYPE_KONSTRUKSI = DAYA_DUKUNG.TYPE_KONSTRUKSI WHERE DBKB_DAYA_DUKUNG.THN_DBKB_DAYA_DUKUNG='" & ccTahun.Text & "'"
QSTR = "Select * FROM DAT_JPB8"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then
                XTIPE = rPajak!TYPE_KONSTRUKSI
                xKOLOM = rPajak!TING_KOLOM_JPB8
                xBENTANG = rPajak!LBR_BENT_JPB8
                xMEZZ = rPajak!LUAS_MEZZANINE_JPB8
                xKELILING = rPajak!KELILING_DINDING_JPB8
                xDUKUNG = rPajak!DAYA_DUKUNG_LANTAI_JPB8
                'If xDUKUNG * 1 >= rPajak!DAYA_DUKUNG_LANTAI_MIN_DBKB And xDUKUNG * 1 <= rPajak!DAYA_DUKUNG_LANTAI_MAX_DBKB Then
                     vBng.ListItems.Item(J).ListSubItems(45).Text = xMEZZ
                     vBng.ListItems.Item(J).ListSubItems(47).Text = xDUKUNG 'rPajak!NILAI_DBKB_DAYA_DUKUNG * vBng.ListItems.Item(j).ListSubItems(13).Text
                     vBng.ListItems.Item(J).ListSubItems(49).Text = xBENTANG
                     vBng.ListItems.Item(J).ListSubItems(50).Text = xKOLOM
                     vBng.ListItems.Item(J).ListSubItems(71).Text = xKELILING
                'End If
            End If
        Next

    
    rPajak.MoveNext
Loop
If i = 0 Then nDUKUNG = 0
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
'============PENILAIAN INDIVIDU==============
Sub NILAI_INDIVIDU1()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_NILAI_INDIVIDU"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        For A = 1 To vOP.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vOP.ListItems.Item(A).ListSubItems(2).Text & "." & _
               vOP.ListItems.Item(A).ListSubItems(3).Text & "." & _
               vOP.ListItems.Item(A).ListSubItems(4).Text & "." & _
               vOP.ListItems.Item(A).ListSubItems(5).Text & "." & _
               vOP.ListItems.Item(A).ListSubItems(6).Text & "." & _
               vOP.ListItems.Item(A).ListSubItems(7).Text & "-" & _
               vOP.ListItems.Item(A).ListSubItems(8).Text '& "." & _
            NOP2 = vOP.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 Then 'And Trim(rPajak!NO_BNG) = vBng.ListItems.Item(A).ListSubItems(9).Text Then
                'vBng.ListItems.Item(j).ListSubItems(26).Text = rPajak!NILAI_INDIVIDU
                vOP.ListItems.Item(A).ListSubItems(14).Text = rPajak!NILAI_INDIVIDU
                vOP.ListItems.Item(A).ListSubItems(18).Text = "INDIVIDU"
            End If
        Next
    rPajak.MoveNext
    Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub NILAI_INDIVIDU()
On Error GoTo Salah
StrQ = "SELECT * FROM DAT_NILAI_INDIVIDU"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        For J = 1 To vBng.ListItems.Count
            NOP1 = Trim(rPajak!KD_PROPINSI) & "." & Trim(rPajak!KD_DATI2) & "." & Trim(rPajak!KD_KECAMATAN) & "." & Trim(rPajak!KD_KELURAHAN) & "." & Trim(rPajak!KD_BLOK) & "-" & Trim(rPajak!NO_URUT) & "." & Trim(rPajak!KD_JNS_OP) '& "." & Trim(rPajak!NO_BNG)
            NOP2 = vBng.ListItems.Item(J).ListSubItems(2).Text & "." & vBng.ListItems.Item(J).ListSubItems(3).Text & "." & vBng.ListItems.Item(J).ListSubItems(4).Text & "." & vBng.ListItems.Item(J).ListSubItems(5).Text & "." & vBng.ListItems.Item(J).ListSubItems(6).Text & "-" & vBng.ListItems.Item(J).ListSubItems(7).Text & "." & vBng.ListItems.Item(J).ListSubItems(8).Text '& "." & vBng.ListItems.Item(J).ListSubItems(9).Text
            If NOP1 = NOP2 And Trim(rPajak!NO_BNG) = vBng.ListItems.Item(J).ListSubItems(9).Text Then
                'vBng.ListItems.Item(j).ListSubItems(26).Text = rPajak!NILAI_INDIVIDU
                vBng.ListItems.Item(J).ListSubItems(80).Text = rPajak!NILAI_INDIVIDU
                vBng.ListItems.Item(J).ListSubItems(69).Text = "INDIVIDU"
            End If
        Next
    rPajak.MoveNext
    Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
'============BANGUNAN TIDAK KENA PAJAK==============
Sub TIDAK_KENA_PAJAK()
On Error GoTo Salah
        For J = 1 To vBng.ListItems.Count
            If Trim(vBng.ListItems.Item(J).ListSubItems(10).Text) = 11 Then
                'vBng.ListItems.Item(j).ListSubItems(26).Text = rPajak!NILAI_INDIVIDU
                vBng.ListItems.Item(J).ListSubItems(63).Text = 0
                'vBng.ListItems.Item(J).ListSubItems(69).Text = "INDIVIDU"
            End If
        Next
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub

'===========MENENTUKAN DBKB MATERIAL==============
Sub MATERIAL()
On Error GoTo Salah
Dim jumLT
StrQ = "SELECT * FROM DBKB_MATERIAL WHERE THN_DBKB_MATERIAL='" & ccTahun.Text & "' ORDER BY KD_PEKERJAAN ASC"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    K = 0
    If rPajak.RecordCount > 1 Then 'vBng.ListItems.Count > 1 Then
        pNilai.Max = rPajak.RecordCount ' * vBng.ListItems.Count
        pNilai.Min = 1
    End If
    tJum.Refresh
    tJum.Text = vBng.ListItems.Count 'rPajak.RecordCount
    Do While Not rPajak.EOF
        K = K + 1
        'If K > vBng.ListItems.Count Then K = 1
           
            LNilai.Refresh
        For J = 1 To vBng.ListItems.Count
           LNilai.Visible = True
            LNilai.Caption = "4/13 - Proses DBKB Material: " & Round(K / pNilai.Max * 100, 0) & "%"  '" & Titik1(k)
            LNilai.Visible = False
            pNilai.Value = K
            If rPajak!KD_PEKERJAAN = "21" Then
                XX1 = vBng.ListItems.Item(J).ListSubItems(18).Text '4
                XX2 = rPajak!KD_KEGIATAN * 1
                If XX1 <= 3 Then
                    XX1 = vBng.ListItems.Item(J).ListSubItems(18).Text * 1
                Else 'If XX1 >= 4 Then
                    XX1 = (vBng.ListItems.Item(J).ListSubItems(18).Text * 1) + 3
                End If
                If XX1 = XX2 Then vBng.ListItems.Item(J).ListSubItems(57).Text = rPajak!NILAI_DBKB_MATERIAL
            End If
            If rPajak!KD_PEKERJAAN = "22" And vBng.ListItems.Item(J).ListSubItems(19).Text = rPajak!KD_KEGIATAN * 1 Then
                vBng.ListItems.Item(J).ListSubItems(58).Text = rPajak!NILAI_DBKB_MATERIAL
            End If
            If rPajak!KD_PEKERJAAN = "23" And vBng.ListItems.Item(J).ListSubItems(17).Text * 1 = rPajak!KD_KEGIATAN * 1 Then
                jumLT = vBng.ListItems.Item(J).ListSubItems(14).Text
                If jumLT = "" Or jumLT = 0 Then
                    jumLT = 1
                End If
                
                vBng.ListItems.Item(J).ListSubItems(56).Text = rPajak!NILAI_DBKB_MATERIAL / jumLT ' & "-" & rPajak!KD_KEGIATAN * 1
            End If
            If rPajak!KD_PEKERJAAN = "24" And vBng.ListItems.Item(J).ListSubItems(20).Text = rPajak!KD_KEGIATAN * 1 Then
                vBng.ListItems.Item(J).ListSubItems(59).Text = rPajak!NILAI_DBKB_MATERIAL
            End If
        xPro(0).Text = J
        
        xPro(1).Text = vBng.ListItems.Item(J).ListSubItems(64).Text
        
    Next
        xPro(0).Refresh
        xPro(1).Refresh
    rPajak.MoveNext
    Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

'=======Menentukan Nilai Fasilitas
Sub DBKB_FAS1A()
On Error GoTo Salah
QSTR = "SELECT FASILITAS.KD_FASILITAS, FASILITAS.NM_FASILITAS, FASILITAS.SATUAN_FASILITAS, FAS_NON_DEP.NILAI_NON_DEP, FAS_NON_DEP.THN_NON_DEP FROM FASILITAS INNER JOIN FAS_NON_DEP ON FASILITAS.KD_FASILITAS = FAS_NON_DEP.KD_FASILITAS WHERE FAS_NON_DEP.THN_NON_DEP='" & ccTahun.Text & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
For J = 1 To vFAS.ListItems.Count
    If vFAS.ListItems.Item(J).ListSubItems(10).Text * 1 = rPajak!KD_FASILITAS * 1 Then
        vFAS.ListItems.Item(J).ListSubItems(12).Text = rPajak!NILAI_NON_DEP
        If vFAS.ListItems.Item(J).ListSubItems(10).Text = "44" Then vFAS.ListItems.Item(J).ListSubItems(12).Text = rPajak!NILAI_NON_DEP / 1000
        If rPajak!KD_FASILITAS = "01" Or rPajak!KD_FASILITAS = "02" Or rPajak!KD_FASILITAS = "11" Or rPajak!KD_FASILITAS = "44" Then  'Tidak Dihitung Susut
            vFAS.ListItems.Item(J).ListSubItems(13).Text = vFAS.ListItems.Item(J).ListSubItems(11).Text * vFAS.ListItems.Item(J).ListSubItems(12).Text
        Else
            'Dihitung Susut
            vFAS.ListItems.Item(J).ListSubItems(15).Text = vFAS.ListItems.Item(J).ListSubItems(11).Text * vFAS.ListItems.Item(J).ListSubItems(12).Text
            vFAS.ListItems.Item(J).ListSubItems(19).Text = vFAS.ListItems.Item(J).ListSubItems(12).Text
        End If
    End If
Next
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub DBKB_FAS3A()
On Error GoTo Salah
QSTR = "SELECT FASILITAS.KD_FASILITAS, FASILITAS.NM_FASILITAS, FASILITAS.SATUAN_FASILITAS, FAS_DEP_MIN_MAX.KLS_DEP_MIN, FAS_DEP_MIN_MAX.KLS_DEP_MAX, FAS_DEP_MIN_MAX.NILAI_DEP_MIN_MAX, FAS_DEP_MIN_MAX.THN_DEP_MIN_MAX FROM FASILITAS INNER JOIN FAS_DEP_MIN_MAX ON FASILITAS.KD_FASILITAS = FAS_DEP_MIN_MAX.KD_FASILITAS WHERE FAS_DEP_MIN_MAX.THN_DEP_MIN_MAX='" & ccTahun.Text & "' ORDER BY FASILITAS.KD_FASILITAS,FAS_DEP_MIN_MAX.KLS_DEP_MIN,FAS_DEP_MIN_MAX.KLS_DEP_MAX ASC"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
For J = 1 To vFAS.ListItems.Count
    If vFAS.ListItems.Item(J).ListSubItems(10).Text * 1 = rPajak!KD_FASILITAS * 1 Then
        If vFAS.ListItems.Item(J).ListSubItems(11).Text * 1 >= rPajak!KLS_DEP_MIN And vFAS.ListItems.Item(J).ListSubItems(11).Text <= rPajak!KLS_DEP_MAX Then
            vFAS.ListItems.Item(J).ListSubItems(12).Text = rPajak!NILAI_DEP_MIN_MAX
            If rPajak!KD_FASILITAS = "40" Then
                vFAS.ListItems.Item(J).ListSubItems(13).Text = vFAS.ListItems.Item(J).ListSubItems(11).Text * vFAS.ListItems.Item(J).ListSubItems(12).Text 'Tidak Susut
            Else
                vFAS.ListItems.Item(J).ListSubItems(15).Text = vFAS.ListItems.Item(J).ListSubItems(11).Text * vFAS.ListItems.Item(J).ListSubItems(12).Text 'Susut
                vFAS.ListItems.Item(J).ListSubItems(19).Text = vFAS.ListItems.Item(J).ListSubItems(12).Text 'Susut
            End If
            'vFAS.ListItems.Item(j).ListSubItems(13).Text = vFAS.ListItems.Item(j).ListSubItems(11).Text * vFAS.ListItems.Item(j).ListSubItems(12).Text

        End If
    End If
Next
rPajak.MoveNext
Loop

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub DBKB_FAS2A()
'Tidak Ada Disusutkan
On Error GoTo Salah
QSTR = "SELECT FASILITAS.KD_FASILITAS, FASILITAS.NM_FASILITAS, FASILITAS.SATUAN_FASILITAS, FAS_DEP_JPB_KLS_BINTANG.KLS_BINTANG, FAS_DEP_JPB_KLS_BINTANG.NILAI_FASILITAS_KLS_BINTANG, FAS_DEP_JPB_KLS_BINTANG.THN_DEP_JPB_KLS_BINTANG FROM FASILITAS INNER JOIN FAS_DEP_JPB_KLS_BINTANG ON FASILITAS.KD_FASILITAS = FAS_DEP_JPB_KLS_BINTANG.KD_FASILITAS WHERE FAS_DEP_JPB_KLS_BINTANG.THN_DEP_JPB_KLS_BINTANG='" & ccTahun.Text & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
For J = 1 To vFAS.ListItems.Count
    If vFAS.ListItems.Item(J).ListSubItems(10).Text * 1 = rPajak!KD_FASILITAS * 1 Then
        If vFAS.ListItems.Item(J).ListSubItems(11).Text * 1 = Trim(rPajak!KLS_BINTANG) * 1 Then
            vFAS.ListItems.Item(J).ListSubItems(12).Text = rPajak!NILAI_FASILITAS_KLS_BINTANG
            If vFAS.ListItems.Item(J).ListSubItems(10).Text = "03" Or vFAS.ListItems.Item(J).ListSubItems(10).Text = "04" Or vFAS.ListItems.Item(J).ListSubItems(10).Text = "06" Or vFAS.ListItems.Item(J).ListSubItems(10).Text = "07" Or vFAS.ListItems.Item(J).ListSubItems(10).Text = "09" Then 'Or vFAS.ListItems.Item(j).ListSubItems(10).Text = "43" Or vFAS.ListItems.Item(j).ListSubItems(10).Text = "45" Then
                vFAS.ListItems.Item(J).ListSubItems(16).Text = vFAS.ListItems.Item(J).ListSubItems(12).Text '*vFAS.ListItems.Item(j).ListSubItems(11).Text
            ElseIf vFAS.ListItems.Item(J).ListSubItems(10).Text = "05" Or vFAS.ListItems.Item(J).ListSubItems(10).Text = "08" Or vFAS.ListItems.Item(J).ListSubItems(10).Text = "10" Then
                vFAS.ListItems.Item(J).ListSubItems(17).Text = vFAS.ListItems.Item(J).ListSubItems(12).Text '*vFAS.ListItems.Item(j).ListSubItems(11).Text
            Else 'If vFAS.ListItems.Item(j).ListSubItems(10).Text = "43" Or vFAS.ListItems.Item(j).ListSubItems(10).Text = "45" Then
                vFAS.ListItems.Item(J).ListSubItems(18).Text = vFAS.ListItems.Item(J).ListSubItems(12).Text '*vFAS.ListItems.Item(j).ListSubItems(11).Text
            End If
        End If
    End If
Next
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

'================PANGGIL NILAI FASILITAS================

Sub NILAI_FASILITAS()
On Error GoTo Salah

Timer1.Enabled = True
vFAS.ListItems.Clear
    'StrQ = "SELECT DAT_OP_BUMI.KD_PROPINSI, DAT_OP_BUMI.KD_DATI2, DAT_OP_BUMI.KD_KECAMATAN, DAT_OP_BUMI.KD_KELURAHAN, DAT_OP_BUMI.KD_BLOK, DAT_OP_BUMI.NO_URUT, DAT_OP_BUMI.KD_JNS_OP, DAT_OP_BUMI.NO_BUMI, DAT_OP_BUMI.KD_ZNT, DAT_NIR.NIR, DAT_OP_BUMI.LUAS_BUMI, DAT_OP_BUMI.JNS_BUMI, DAT_OP_BUMI.NILAI_SISTEM_BUMI, DAT_NIR.[THN_NIR_ZNT]" & _
            "FROM DAT_NIR INNER JOIN DAT_OP_BUMI ON (DAT_NIR.KD_ZNT = DAT_OP_BUMI.KD_ZNT) AND (DAT_NIR.KD_KELURAHAN = DAT_OP_BUMI.KD_KELURAHAN) AND (DAT_NIR.KD_KECAMATAN = DAT_OP_BUMI.KD_KECAMATAN) WHERE DAT_NIR.[THN_NIR_ZNT]='" & ccTahun.Text & "' AND DAT_OP_BUMI.KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND DAT_OP_BUMI.KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'"
    If ccProses = 1 Then
        StrQ = "SELECT * FROM DAT_FASILITAS_BANGUNAN "
    ElseIf ccProses = 2 Then
        StrQ = "SELECT * FROM DAT_FASILITAS_BANGUNAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' "
    Else
        StrQ = "SELECT * FROM DAT_FASILITAS_BANGUNAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'"
    End If
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    frmBar.Show
'    frmBar.Bar1.Max = rPajak.RecordCount
'    frmBar.Bar1.Min = 1
        If rPajak.RecordCount > 1 Then
            pNilai.Max = rPajak.RecordCount
            pNilai.Min = 1
        End If
        tJum.Refresh
        tJum.Text = rPajak.RecordCount
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        LNilai.Caption = "5/13 - Proses Penilaian Fasilitas: " & Round(i / pNilai.Max * 100, 0) & "%" '" & Titik2(j)
        pNilai.Value = i
        LNilai.Refresh
        LNilai.Visible = False

        vFAS.ListItems.Add i, "", Format(i, "#")
        vFAS.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#,#0")
        xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
        xPro(0).Text = i
        xPro(0).Refresh
        xPro(1).Text = xNOP
        xPro(1).Refresh
        vFAS.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_PROPINSI])
        vFAS.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![KD_DATI2])
        vFAS.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![KD_KECAMATAN])
        vFAS.ListItems.Item(i).ListSubItems.Add 5, "", Trim(rPajak![KD_KELURAHAN])
        vFAS.ListItems.Item(i).ListSubItems.Add 6, "", Trim(rPajak![KD_BLOK])
        vFAS.ListItems.Item(i).ListSubItems.Add 7, "", Trim(rPajak![NO_URUT])
        vFAS.ListItems.Item(i).ListSubItems.Add 8, "", Trim(rPajak![KD_JNS_OP])
        vFAS.ListItems.Item(i).ListSubItems.Add 9, "", Trim(rPajak![NO_BNG])
        vFAS.ListItems.Item(i).ListSubItems.Add 10, "", Trim(rPajak![KD_FASILITAS])
        vFAS.ListItems.Item(i).ListSubItems.Add 11, "", Trim(rPajak![JML_SATUAN])
        '------------
        
        vFAS.ListItems.Item(i).ListSubItems.Add 12, "", 0 'xTahun
        vFAS.ListItems.Item(i).ListSubItems.Add 13, "", 0 'Trim(rPajak![LUAS_BNG])
        vFAS.ListItems.Item(i).ListSubItems.Add 14, "", 0 'Trim(rPajak![JML_LANTAI_BNG])
        vFAS.ListItems.Item(i).ListSubItems.Add 15, "", 0 'Trim(rPajak![KONDISI_BNG])
        vFAS.ListItems.Item(i).ListSubItems.Add 16, "", 0 'Trim(rPajak![JNS_KONSTRUKSI_BNG])
        vFAS.ListItems.Item(i).ListSubItems.Add 17, "", 0 'Trim(rPajak![JNS_ATAP_BNG])
        vFAS.ListItems.Item(i).ListSubItems.Add 18, "", 0 'Trim(rPajak![KD_DINDING])
        vFAS.ListItems.Item(i).ListSubItems.Add 19, "", 0 'NIlai DBKB Fasilitas
        vFAS.ListItems.Item(i).ListSubItems.Add 20, "", 0 'Trim(rPajak![KD_LANGIT_LANGIT])
        vFAS.ListItems.Item(i).ListSubItems.Add 21, "", 0 'Format(rPajak![NILAI_SISTEM_BNG], "#,#0.00")
        vFAS.ListItems.Item(i).ListSubItems.Add 22, "", 0 'Trim(rPajak![JNS_BUMI])
        vFAS.ListItems.Item(i).ListSubItems.Add 23, "", 0 'Trim(rPajak![JNS_BUMI])
        vFAS.ListItems.Item(i).ListSubItems.Add 24, "", 0 'Format(rPajak![NILAI_SISTEM_BNG] / rPajak![LUAS_BNG], "#,#0.00")
        vFAS.ListItems.Item(i).ListSubItems.Add 25, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 26, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 27, "", 0
        '=======
        vFAS.ListItems.Item(i).ListSubItems.Add 28, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 29, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 30, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 31, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 32, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 33, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 34, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 35, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 36, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 37, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 38, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 39, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 40, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 41, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 42, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 43, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 44, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 45, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 46, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 47, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 48, "", 0
        '===========
        vFAS.ListItems.Item(i).ListSubItems.Add 49, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 50, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 51, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 52, "", 0
        '-------------
        vFAS.ListItems.Item(i).ListSubItems.Add 53, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 54, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 55, "", 0
        '==================
        vFAS.ListItems.Item(i).ListSubItems.Add 56, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 57, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 58, "", 0
        vFAS.ListItems.Item(i).ListSubItems.Add 59, "", 0
        


'                frmBar.Bar1.Value = i
    rPajak.MoveNext
    Loop
call_judul_fas
'SALAH:
'If Err.Number = 0 Then Exit Sub
'MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error"

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
'=============Jumlahkan Nilai Fasilitas===============
Sub NFAS()
On Error Resume Next 'On Error GoTo Salah
Dim cJum, nMaterial, nMez_DDukung, cJum1, cJum2, cJum3, cJum4, s_FJ


For J = 1 To vFAS.ListItems.Count
        NOP1 = vFAS.ListItems.Item(J).ListSubItems(2).Text & "." & vFAS.ListItems.Item(J).ListSubItems(3).Text & "." & vFAS.ListItems.Item(J).ListSubItems(4).Text & "." & vFAS.ListItems.Item(J).ListSubItems(5).Text & "." & vFAS.ListItems.Item(J).ListSubItems(6).Text & "-" & vFAS.ListItems.Item(J).ListSubItems(7).Text & "." & vFAS.ListItems.Item(J).ListSubItems(8).Text '& "." & vFAS.ListItems.Item(J).ListSubItems(9).Text
        vFAS.ListItems.Item(J).ListSubItems(14).Text = NOP1
        
Next
If vBng.ListItems.Count > 1 Then
            pNilai.Max = vBng.ListItems.Count
            pNilai.Min = 1
        End If
        tJum.Refresh
tJum.Text = vBng.ListItems.Count
For K = 1 To vBng.ListItems.Count
            LNilai.Visible = True
            LNilai.Caption = "5/13 - Proses Nilai Fasilitas: " & Round(K / pNilai.Max * 100, 0) & "%" '" & Titik1(k)
            LNilai.Refresh
            LNilai.Visible = False
            pNilai.Value = K
        cJum = 0: nMaterial = 0: cJum1 = 0: cJum2 = 0: cJum3 = 0: cJum4 = 0
        s_FJ = 0
        NOP2 = vBng.ListItems.Item(K).ListSubItems(2).Text & "." & vBng.ListItems.Item(K).ListSubItems(3).Text & "." & vBng.ListItems.Item(K).ListSubItems(4).Text & "." & vBng.ListItems.Item(K).ListSubItems(5).Text & "." & vBng.ListItems.Item(K).ListSubItems(6).Text & "-" & vBng.ListItems.Item(K).ListSubItems(7).Text & "." & vBng.ListItems.Item(K).ListSubItems(8).Text '& "." & vBng.ListItems.Item(K).ListSubItems(9).Text
        vBng.ListItems.Item(K).ListSubItems(64).Text = NOP2
        For L = 1 To vFAS.ListItems.Count
            If (vBng.ListItems.Item(K).ListSubItems(64).Text = vFAS.ListItems.Item(L).ListSubItems(14).Text) And (vBng.ListItems.Item(K).ListSubItems(9).Text = vFAS.ListItems.Item(L).ListSubItems(9).Text) Then
                cJum = cJum + (vFAS.ListItems.Item(L).ListSubItems(13).Text * 1)
                cJum1 = cJum1 + (vFAS.ListItems.Item(L).ListSubItems(15).Text * 1)
                cJum2 = cJum2 + (vFAS.ListItems.Item(L).ListSubItems(16).Text * 1)
                cJum3 = cJum3 + (vFAS.ListItems.Item(L).ListSubItems(17).Text * 1)
                cJum4 = cJum4 + (vFAS.ListItems.Item(L).ListSubItems(18).Text * 1)
                s_FJ = s_FJ + (vFAS.ListItems.Item(L).ListSubItems(19).Text * 1)
            End If
        Next
        vBng.ListItems.Item(K).ListSubItems(60).Text = cJum
        vBng.ListItems.Item(K).ListSubItems(61).Text = cJum1
        vBng.ListItems.Item(K).ListSubItems(66).Text = cJum2
        vBng.ListItems.Item(K).ListSubItems(67).Text = cJum3
        vBng.ListItems.Item(K).ListSubItems(68).Text = cJum4
        vBng.ListItems.Item(K).ListSubItems(70).Text = s_FJ
        
        
Next

vFAS.SortKey = 14
vFAS.Sorted = True
vFAS.Sorted = False
vFAS.SortOrder = lvwAscending

vBng.SortKey = 61
vBng.Sorted = True
vBng.Sorted = False
vBng.SortOrder = lvwAscending
Call_Susut
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'Salah:
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'MsgBox Err.Number & ": " & Err.Description
End Sub
Sub Call_Susut()
On Error Resume Next 'GoTo Salah
Dim xUmum
Dim TPajak, TBangun, TRenovasi, JGuna, JLANTAI, umur_EFF
Dim n_Fas1, n_Fas2, nSusut
Dim JPB1, JPB2, JPB3, JPB4, JPB5, JPB6, JPB7, JPB8, JPB9, JPB10, JPB11, JPB12, JPB13, JPB14, JPB15, JPB16, JPB17
Dim xSistem1, xSistem2, nAC, nBoiler
Dim nMateral1, nMaterial2
'xKondisi = vBng.ListItems.Item(A).ListSubItems(15).Text

StrQ = "SELECT PENYUSUTAN.KD_RANGE_PENYUSUTAN, PENYUSUTAN.UMUR_EFEKTIF, PENYUSUTAN.KONDISI_BNG_SUSUT, RANGE_PENYUSUTAN.NILAI_MIN_PENYUSUTAN, RANGE_PENYUSUTAN.NILAI_MAX_PENYUSUTAN, PENYUSUTAN.NILAI_PENYUSUTAN FROM PENYUSUTAN INNER JOIN RANGE_PENYUSUTAN ON PENYUSUTAN.KD_RANGE_PENYUSUTAN = RANGE_PENYUSUTAN.KD_RANGE_PENYUSUTAN ORDER BY PENYUSUTAN.UMUR_EFEKTIF" 'where PENYUSUTAN.KONDISI_BNG_SUSUT*1='" & xKondisi & "' and PENYUSUTAN.UMUR_EFEKTIF='" & umur_EFF & "' ORDER BY PENYUSUTAN.UMUR_EFEKTIF"
openDB (StrQ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
        If vBng.ListItems.Count > 1 Then
            pNilai.Max = rPajak.RecordCount ' vBng.ListItems.Count
            pNilai.Min = 1
        End If
        tJum.Refresh
tJum.Text = rPajak.RecordCount
Do While Not rPajak.EOF
        i = i + 1
        
    LNilai.Refresh
    For J = 1 To vBng.ListItems.Count
    'i = 0: J = 0
        LNilai.Visible = True
        LNilai.Caption = "6/13 - Proses Total Nilai : " & Round(i / pNilai.Max * 100, 0) & "%" '" & Titik2(k)
        pNilai.Value = i
        
        'LNilai.Visible = False
        'Do While Not rPajak.EOF
        
    JGuna = vBng.ListItems.Item(J).ListSubItems(10).Text
     'Material
     If JGuna = 3 Then
        nMaterial1 = ((vBng.ListItems.Item(J).ListSubItems(59).Text * 1) + (vBng.ListItems.Item(J).ListSubItems(58).Text * 1) + (vBng.ListItems.Item(J).ListSubItems(56).Text * 1)) * 1.3 * vBng.ListItems.Item(J).ListSubItems(13).Text
        nMaterial2 = vBng.ListItems.Item(J).ListSubItems(13).Text * vBng.ListItems.Item(J).ListSubItems(71).Text * vBng.ListItems.Item(J).ListSubItems(50).Text * (10 / 6) * vBng.ListItems.Item(J).ListSubItems(57).Text * 1.3 'Keliling*Tinggi Kolom*10/6*dinding
        nMaterial = nMateral1 + nMaterial2
     ElseIf JGuna = 8 Then
        nMaterial1 = ((vBng.ListItems.Item(J).ListSubItems(59).Text * 1) + (vBng.ListItems.Item(J).ListSubItems(58).Text * 1) + (vBng.ListItems.Item(J).ListSubItems(56).Text * 1)) * vBng.ListItems.Item(J).ListSubItems(13).Text
        nMaterial2 = vBng.ListItems.Item(J).ListSubItems(13).Text * vBng.ListItems.Item(J).ListSubItems(71).Text * vBng.ListItems.Item(J).ListSubItems(50).Text * (10 / 6) * vBng.ListItems.Item(J).ListSubItems(57).Text
        nMaterial = nMateral1 + nMaterial2
     Else
        nMaterial = ((vBng.ListItems.Item(J).ListSubItems(59).Text * 1) + (vBng.ListItems.Item(J).ListSubItems(58).Text * 1) + (vBng.ListItems.Item(J).ListSubItems(57).Text * 1) + (vBng.ListItems.Item(J).ListSubItems(56).Text * 1)) * vBng.ListItems.Item(J).ListSubItems(13).Text
     End If
    
    'nMezanin dan Daya Dukung
    nMez_DDukung = vBng.ListItems.Item(J).ListSubItems(46).Text + vBng.ListItems.Item(J).ListSubItems(48).Text
    
    'Fasilitas
    If vBng.ListItems.Item(J).ListSubItems(10).Text = "02" Or vBng.ListItems.Item(J).ListSubItems(10).Text = "04" Then
        nAC = (vBng.ListItems.Item(J).ListSubItems(66).Text * vBng.ListItems.Item(J).ListSubItems(13).Text) ' + (vBng.ListItems.Item(j).ListSubItems(67).Text * vBng.ListItems.Item(j).ListSubItems(13).Text)
    Else
        nAC = (vBng.ListItems.Item(J).ListSubItems(66).Text * vBng.ListItems.Item(J).ListSubItems(54).Text) + (vBng.ListItems.Item(J).ListSubItems(67).Text * vBng.ListItems.Item(J).ListSubItems(55).Text)
    End If
    'vBng.ListItems.Item(j).ListSubItems(68).Text = nAC
    nBoiler = vBng.ListItems.Item(J).ListSubItems(68).Text * vBng.ListItems.Item(J).ListSubItems(53).Text
  '  nFasilitas = vBng.ListItems.Item(J).ListSubItems(60).Text
    'DBKB Per Jenis Bangunan
    
    'JPB=1
        JPB1 = vBng.ListItems.Item(J).ListSubItems(28).Text * 1
    'JPB=2
        If vBng.ListItems.Item(J).ListSubItems(29).ForeColor = vbRed Then
            vBng.ListItems.Item(J).ListSubItems(29).Text = 0
        End If
        JPB2 = vBng.ListItems.Item(J).ListSubItems(29).Text
    'JPB=3
        JPB3 = vBng.ListItems.Item(J).ListSubItems(30).Text * 1
    'JPB=4
        If vBng.ListItems.Item(J).ListSubItems(31).ForeColor = vbRed Then
            vBng.ListItems.Item(J).ListSubItems(31).Text = 0
        End If
        JPB4 = vBng.ListItems.Item(J).ListSubItems(31).Text * 1
    'JPB=5
        If vBng.ListItems.Item(J).ListSubItems(32).ForeColor = vbRed Then
            vBng.ListItems.Item(J).ListSubItems(32).Text = 0
        End If
        JPB5 = vBng.ListItems.Item(J).ListSubItems(32).Text * 1
    'JPB=6
        JPB6 = vBng.ListItems.Item(J).ListSubItems(33).Text * 1
    'JPB=7
        If vBng.ListItems.Item(J).ListSubItems(34).ForeColor = vbRed Then
            vBng.ListItems.Item(J).ListSubItems(34).Text = 0
        End If
            JPB7 = vBng.ListItems.Item(J).ListSubItems(34).Text * 1
    'JPB=8
        JPB8 = vBng.ListItems.Item(J).ListSubItems(35).Text * 1
    'JPB=9
        If vBng.ListItems.Item(J).ListSubItems(36).ForeColor = vbRed Then
            vBng.ListItems.Item(J).ListSubItems(36).Text = 0
        End If
            JPB9 = vBng.ListItems.Item(J).ListSubItems(36).Text * 1
    'JPB=10
        JPB10 = vBng.ListItems.Item(J).ListSubItems(37).Text * 1
    'JPB=11
        JPB11 = vBng.ListItems.Item(J).ListSubItems(38).Text * 1
    'JPB=12
        JPB12 = vBng.ListItems.Item(J).ListSubItems(39).Text * 1
    'JPB=13
        JPB13 = vBng.ListItems.Item(J).ListSubItems(40).Text * 1
    'JPB=14
        JPB14 = vBng.ListItems.Item(J).ListSubItems(41).Text * 1
    'JPB=15
        JPB15 = vBng.ListItems.Item(J).ListSubItems(42).Text * 1
    'JPB=16
        JPB16 = vBng.ListItems.Item(J).ListSubItems(43).Text * 1
    'JPB=17
        JPB17 = vBng.ListItems.Item(J).ListSubItems(44).Text * 1
    
    nSistem_B4_Susut = (JPB1 + JPB2 + JPB3 + JPB4 + JPB5 + JPB6 + JPB7 + JPB8 + JPB9 + JPB10 + JPB11 + JPB12 + JPB13 + JPB14 + JPB15 + JPB16 + JPB17) * vBng.ListItems.Item(J).ListSubItems(13).Text
    
    vBng.ListItems.Item(J).ListSubItems(65).Text = (JPB1 + JPB2 + JPB3 + JPB4 + JPB5 + JPB6 + JPB7 + JPB8 + JPB9 + JPB10 + JPB11 + JPB12 + JPB13 + JPB14 + JPB15 + JPB16 + JPB17)
    'Menentukan Nilai Penyusutan
                TPajak = ccTahun.Text * 1
                TBangun = vBng.ListItems.Item(J).ListSubItems(11).Text
                TRenovasi = vBng.ListItems.Item(J).ListSubItems(12).Text
                'JGuna = vBng.ListItems.Item(J).ListSubItems(10).Text
                JLANTAI = vBng.ListItems.Item(J).ListSubItems(14).Text
                'If TRenovasi = "-" Or TRenovasi = "" Then TRenovasi = 0
                'Mencari Umur Efektif
                '-------------------------------------
                'Bangunan Standar
                '-------------------------------------
                If (JGuna = 1 Or JGuna = 3 Or JGuna = 8 Or JGuna = 10 Or JGuna = 11 Or JGuna = 2 Or JGuna = 4 Or JGuna = 5 Or JGuna = 7 Or JGuna = 9) And JLANTAI <= 4 Then
                    If TRenovasi > 0 Then
                        Umur = TPajak - TRenovasi
                    Else
                        Umur = TPajak - TBangun
                    End If
                '-------------------------------------
                'Bangunan Non Standar
                '-------------------------------------
                Else
                    If TBangun > 0 And TRenovasi > 0 Then
                        If TPajak - TBangun > 10 Then
                            Umur = ((TPajak - TBangun) + (2 * 10)) / 3
                        Else
                            Umur = ((TPajak - TBangun) + 2 * (TPajak - TRenovasi)) / 3
                        End If
                    Else 'If TBangun > 0 And TRenovasi <= 0 Then
                        If TPajak - TBangun > 10 Then
                            Umur = ((TPajak - TBangun) + (2 * 10)) / 3
                        Else
                            Umur = TPajak - TBangun
                        End If
                    End If
                    nMaterial = 0
                End If

                
                
                
'                If TRenovasi <= 0 Then 'Tidak Ada Renovasi
'                    If (JGuna = 1 Or JGuna = 3 Or JGuna = 8 Or JGuna = 10 Or JGuna = 11 Or JGuna = 2 Or JGuna = 4 Or JGuna = 5 Or JGuna = 7 Or JGuna = 9) And JLANTAI <= 4 Then
'                        Umur = TPajak - TBangun
'                    Else
'                        If (TPajak - TBangun) <= 10 Then
'                            Umur = TPajak - TBangun
'                        Else
'                            Umur = (TPajak - TBangun + 20) / 3
'                        End If
'                    End If
'                Else 'Ada Renovasi
'                    If (JGuna = 1 Or JGuna = 3 Or JGuna = 8 Or JGuna = 10 Or JGuna = 11 Or JGuna = 2 Or JGuna = 4 Or JGuna = 5 Or JGuna = 7 Or JGuna = 9) And JLANTAI <= 4 Then
'                        Umur = TPajak - TRenovasi
'                    Else
'                        If (TRenovasi - TBangun) <= 10 Then
'                            Umur = ((TPajak - TBangun) + (2 * (TPajak - TBangun))) / 3
'                        Else
'                            Umur = (TPajak - TBangun + 20) / 3
'                        End If
'                    End If
'                End If
                
                If Umur > 40 Then
                    Umur = 40
                End If
                umur_EFF = Round(Umur) ' + 0.4)

'
'        Else
        
'        If vBng.ListItems.Item(J).ListSubItems(16).Text * 1 = 4 And (vBng.ListItems.Item(J).ListSubItems(18).Text * 1 = 4 Or vBng.ListItems.Item(J).ListSubItems(18).Text * 1 = 5) Then  'And (vBng.ListItems.Item(J).ListSubItems(20).Text * 1 = 3) Then
'
'             'Khusus Konstruksi Kayu, Dinding=Kayu/Seng dan Langit2 Ada
'             If vBng.ListItems.Item(J).ListSubItems(18).Text * 1 = 5 And vBng.ListItems.Item(J).ListSubItems(20).Text * 1 <> 3 Then
'                If nSistem_B4_Susut * 1000 >= rPajak![NILAI_MIN_PENYUSUTAN] And nSistem_B4_Susut * 1000 <= rPajak![NILAI_MAX_PENYUSUTAN] Then
'                 If (vBng.ListItems.Item(J).ListSubItems(15).Text = rPajak!KONDISI_BNG_SUSUT) And (umur_EFF = rPajak!UMUR_EFEKTIF) Then   '='" & umur_EFF & "' ORDER BY PENYUSUTAN.UMUR_EFEKTIF"
'                      vBng.ListItems.Item(J).ListSubItems(62).Text = rPajak![NILAI_PENYUSUTAN]
'                 End If
'                End If
'            Else 'Khusus Konstruksi Kayu, Dinding=Kayu/Seng dan Langit2 Tidak Ada
'                If vBng.ListItems.Item(J).ListSubItems(65).Text * 1000 >= rPajak![NILAI_MIN_PENYUSUTAN] And vBng.ListItems.Item(J).ListSubItems(65).Text * 1000 <= rPajak![NILAI_MAX_PENYUSUTAN] Then
'                 If (vBng.ListItems.Item(J).ListSubItems(15).Text = rPajak!KONDISI_BNG_SUSUT) And (umur_EFF = rPajak!UMUR_EFEKTIF) Then   '='" & umur_EFF & "' ORDER BY PENYUSUTAN.UMUR_EFEKTIF"
'                      vBng.ListItems.Item(J).ListSubItems(62).Text = rPajak![NILAI_PENYUSUTAN]
'                 End If
'                End If
'             End If
'        Else 'Konstuksi Semi Kayu dan Bukan Kayu
    
        
        
'        If (vBng.ListItems.Item(J).ListSubItems(16).Text * 1) = 4 Then
'            If vBng.ListItems.Item(J).ListSubItems(45).Text = 0 Then xxMez = 0 Else xxMez = vBng.ListItems.Item(J).ListSubItems(46).Text / vBng.ListItems.Item(J).ListSubItems(45).Text
'            xxDD = vBng.ListItems.Item(J).ListSubItems(48).Text
'            nBARU = ((nMaterial + xxDD + (nSistem_B4_Susut * 0.7) + n_Fas2) / vBng.ListItems.Item(J).ListSubItems(13).Text) + xxMez 'Nilai Pembuatan Baru Sebelum Susut
'        Else
'            If vBng.ListItems.Item(J).ListSubItems(45).Text = 0 Or vBng.ListItems.Item(J).ListSubItems(45).Text = "" Then xxMez = 0 Else xxMez = vBng.ListItems.Item(J).ListSubItems(46).Text / vBng.ListItems.Item(J).ListSubItems(45).Text
'            xxDD = vBng.ListItems.Item(J).ListSubItems(48).Text
'            nBARU = ((nMaterial + xxDD + nSistem_B4_Susut + n_Fas2) / vBng.ListItems.Item(J).ListSubItems(13).Text) + xxMez  'Nilai Pembuatan Baru Sebelum Susut
'        End If
            
        'End If
'Jika KOnsturksi adalah kayu
        If (vBng.ListItems.Item(J).ListSubItems(16).Text * 1) = 4 And ck_Ulin = 0 Then
            nSistem_B4_Susut = nSistem_B4_Susut * 0.7
        End If
        n_Fas1 = (vBng.ListItems.Item(J).ListSubItems(60).Text) 'Tidak Susut (Sudah Dikalikan Luas dan DBKB)
        n_Fas2 = (vBng.ListItems.Item(J).ListSubItems(61).Text) 'Susut
        s_Fas2 = (vBng.ListItems.Item(J).ListSubItems(70).Text)  'Pure Hanya DBKB Fasilitas untuk Penyusutan
        'vBng.ListItems.Item(J).ListSubItems(61).Text = nMaterial & ":" & nMez_DDukung
        If vBng.ListItems.Item(J).ListSubItems(45).Text = 0 Or vBng.ListItems.Item(J).ListSubItems(45).Text = "" Then xxMez = 0 Else xxMez = vBng.ListItems.Item(J).ListSubItems(46).Text / vBng.ListItems.Item(J).ListSubItems(45).Text
        xxDD = vBng.ListItems.Item(J).ListSubItems(48).Text
        
        'Bangunan Non Standard
        
        nBaru = ((nMaterial + xxDD + nSistem_B4_Susut) / vBng.ListItems.Item(J).ListSubItems(13).Text) + xxMez + s_Fas2 'Nilai Pembuatan Baru Sebelum Susut
        'Bangunan Standard
        
        If nBaru * 1000 >= rPajak![NILAI_MIN_PENYUSUTAN] And nBaru * 1000 <= rPajak![NILAI_MAX_PENYUSUTAN] Then
                 If (vBng.ListItems.Item(J).ListSubItems(15).Text = rPajak!KONDISI_BNG_SUSUT) And (umur_EFF = rPajak!UMUR_EFEKTIF) Then   '='" & umur_EFF & "' ORDER BY PENYUSUTAN.UMUR_EFEKTIF"
                      vBng.ListItems.Item(J).ListSubItems(62).Text = rPajak![NILAI_PENYUSUTAN]
                 End If
        End If
        'Batasi Nilai Peyusutan Untuk JPB=15 (Tangki Minyak)
        If nSusut > 50 And vBng.ListItems.Item(J).ListSubItems(10).Text = "15" Then
            nSusut = 50
        End If
        nSusut = (vBng.ListItems.Item(J).ListSubItems(62).Text) 'Hasil Persentase Susut
        'vBng.ListItems.Item(j).ListSubItems(61).Text = nAC + nBoiler
        xSistem1 = nMaterial + nMez_DDukung + nSistem_B4_Susut + n_Fas2 + n_Fas1 + nAC + nBoiler
        xSistem2 = (nMaterial + xxDD + nSistem_B4_Susut + n_Fas2) * (nSusut / 100) ' + (n_Fas1 + nAC + nBoiler)
        vBng.ListItems.Item(J).ListSubItems(63).Text = Format(xSistem1 - xSistem2, "#,#0.00")
        vBng.ListItems.Item(J).ListSubItems(81).Text = nMez_DDukung + nSistem_B4_Susut 'KOMPONEN UTAMA
        vBng.ListItems.Item(J).ListSubItems(82).Text = nMaterial 'MATERIAL
        vBng.ListItems.Item(J).ListSubItems(83).Text = n_Fas2 'FASILITAS YANG DISUSUTKAN
        vBng.ListItems.Item(J).ListSubItems(84).Text = nSusut 'Jumlah Persentasi Susut
        vBng.ListItems.Item(J).ListSubItems(85).Text = xSistem2 'Nilai Penyusutan
        vBng.ListItems.Item(J).ListSubItems(86).Text = n_Fas1 + nAC + nBoiler 'Fasilitas Yang Tidak Disusutkan
    Next
        xPro(0).Text = i
        xPro(0).Refresh
        xPro(1).Text = vBng.ListItems.Item(i).ListSubItems(64).Text
        xPro(1).Refresh
rPajak.MoveNext
Loop
    LNilai.Visible = True
    LNilai.Caption = "Sukses!"
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'Salah:
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'MsgBox Err.Number & ": " & Err.Description

End Sub

Sub CALL_OP()
'On Error GoTo Salah
Dim QJum1, QJum2, QJum3, rCount

Timer1.Enabled = True
'frmBar.Show
'frmBar.Bar1.Max = vBangunan.ListItems.Count
'frmBar.Bar1.Min = 1
callTarif
J = 0
        If vBangunan.ListItems.Count > 1 Then
            pNilai.Max = vBangunan.ListItems.Count
            pNilai.Min = 1
        End If
tJum.Refresh
tJum.Text = vBangunan.ListItems.Count
call_judul_op
For A = 1 To vBangunan.ListItems.Count
        LNilai.Visible = True
        LNilai.Caption = "7/13 - Proses Pembentukan Objek Pajak: " & Round(A / pNilai.Max * 100, 0) & "%" '" & Titik3(j)
        pNilai.Value = A
        LNilai.Refresh
        LNilai.Visible = False
        vOP.ListItems.Add A, "", Format(A, "#")
        vOP.ListItems.Item(A).ListSubItems.Add 1, "", Format(A, "#,#0")
        'xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
        vOP.ListItems.Item(A).ListSubItems.Add 2, "", vBangunan.ListItems.Item(A).ListSubItems(2).Text
        vOP.ListItems.Item(A).ListSubItems.Add 3, "", vBangunan.ListItems.Item(A).ListSubItems(3).Text
        vOP.ListItems.Item(A).ListSubItems.Add 4, "", vBangunan.ListItems.Item(A).ListSubItems(4).Text
        vOP.ListItems.Item(A).ListSubItems.Add 5, "", vBangunan.ListItems.Item(A).ListSubItems(5).Text
        vOP.ListItems.Item(A).ListSubItems.Add 6, "", vBangunan.ListItems.Item(A).ListSubItems(6).Text
        vOP.ListItems.Item(A).ListSubItems.Add 7, "", vBangunan.ListItems.Item(A).ListSubItems(7).Text
        vOP.ListItems.Item(A).ListSubItems.Add 8, "", vBangunan.ListItems.Item(A).ListSubItems(8).Text
        vOP.ListItems.Item(A).ListSubItems.Add 9, "", vBangunan.ListItems.Item(A).ListSubItems(17).Text 'vBangunan.ListItems.Item(A).ListSubItems(9).Text
        vOP.ListItems.Item(A).ListSubItems.Add 10, "", 0
        vOP.ListItems.Item(A).ListSubItems.Add 11, "", 0
        vOP.ListItems.Item(A).ListSubItems.Add 12, "", 0
        vOP.ListItems.Item(A).ListSubItems.Add 13, "", 0
        vOP.ListItems.Item(A).ListSubItems.Add 14, "", 0
        vOP.ListItems.Item(A).ListSubItems.Add 15, "", 0
        vOP.ListItems.Item(A).ListSubItems.Add 16, "", 0
        vOP.ListItems.Item(A).ListSubItems.Add 17, "", 0
        vOP.ListItems.Item(A).ListSubItems.Add 18, "", "SISTEM"
        vOP.ListItems.Item(A).ListSubItems.Add 19, "", 0
        vOP.ListItems.Item(A).ListSubItems.Add 20, "", vBangunan.ListItems.Item(A).ListSubItems(13).Text 'Jenis Bumi
        vOP.ListItems.Item(A).ListSubItems.Add 21, "", 0 'Nama Wajib Pajak
        vOP.ListItems.Item(A).ListSubItems.Add 22, "", 0 '
        vOP.ListItems.Item(A).ListSubItems.Add 23, "", 0 'Nama Wajib Pajak
        vOP.ListItems.Item(A).ListSubItems.Add 24, "", 0 'Nama Wajib Pajak
        vOP.ListItems.Item(A).ListSubItems.Add 25, "", 0 'Nama Wajib Pajak
        vOP.ListItems.Item(A).ListSubItems.Add 26, "", 0 'Nama Wajib Pajak
        vOP.ListItems.Item(A).ListSubItems.Add 27, "", 0 'Nama Wajib Pajak
        vOP.ListItems.Item(A).ListSubItems.Add 28, "", 0 'Nama Wajib Pajak
        vOP.ListItems.Item(A).ListSubItems.Add 29, "", 0 'Nama Wajib Pajak
        vOP.ListItems.Item(A).ListSubItems.Add 30, "", 0 'Nama Wajib Pajak
        vOP.ListItems.Item(A).ListSubItems.Add 31, "", 0 'Total NJOP : Bumi + Bangunan
        vOP.ListItems.Item(A).ListSubItems.Add 32, "", 0 'Tarif
        vOP.ListItems.Item(A).ListSubItems.Add 33, "", 0 'NJOPTKP
        vOP.ListItems.Item(A).ListSubItems.Add 34, "", 0 'PBB Min
        vOP.ListItems.Item(A).ListSubItems.Add 35, "", 0 'PBB Terutang
        vOP.ListItems.Item(A).ListSubItems.Add 36, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 37, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 38, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 39, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 40, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 41, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 42, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 43, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 44, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 45, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 46, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 47, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 48, "", 0 'ID
        vOP.ListItems.Item(A).ListSubItems.Add 49, "", 0 'ID
'        vOP.ListItems.Item(A).ListSubItems.Add 20, "", rPajak!SUBJEK_PAJAK_ID 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 21, "", rPajak!NO_FORMULIR_SPOP 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 22, "", rPajak!NO_PERSIL 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 23, "", rPajak!JALAN_OP 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 24, "", rPajak!BLOK_KAV_NO_OP 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 25, "", rPajak!RW_OP 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 26, "", rPajak!RT_OP 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 27, "", 0 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 28, "", rPajak!KD_STATUS_WP 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 29, "", rPajak!TOTAL_LUAS_BUMI 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 30, "", rPajak!TOTAL_LUAS_BNG 'Nomor Formulir
'        vOP.ListItems.Item(A).ListSubItems.Add 31, "", rPajak!NJOP_BUMI
'        vOP.ListItems.Item(A).ListSubItems.Add 32, "", rPajak!NJOP_BNG
'        vOP.ListItems.Item(A).ListSubItems.Add 33, "", 1
'        vOP.ListItems.Item(A).ListSubItems.Add 34, "", rPajak!JNS_TRANSAKSI_OP 'Jenis Transaksi
'        vOP.ListItems.Item(A).ListSubItems.Add 35, "", rPajak![TGL_PENDATAAN_OP] 'Tanggal Pendataan
'        vOP.ListItems.Item(A).ListSubItems.Add 36, "", rPajak![NIP_PENDATA] 'NIP Petugas Pendata
'        vOP.ListItems.Item(A).ListSubItems.Add 37, "", rPajak![TGL_PEMERIKSAAN_OP] 'Tanggal Pemeriksaan
'        vOP.ListItems.Item(A).ListSubItems.Add 38, "", rPajak![NIP_PEMERIKSA_OP] 'NIP Petugas Pemeriksa
'        vOP.ListItems.Item(A).ListSubItems.Add 39, "", rPajak![TGL_PEREKAMAN_OP]  'Tanggal Perekaman
'        vOP.ListItems.Item(A).ListSubItems.Add 40, "", rPajak![NIP_PEREKAM_OP] 'NIP Petugas Perekam

        
        NOP1 = vOP.ListItems.Item(A).ListSubItems(2).Text & "." & _
               vOP.ListItems.Item(A).ListSubItems(3).Text & "." & _
               vOP.ListItems.Item(A).ListSubItems(4).Text & "." & _
               vOP.ListItems.Item(A).ListSubItems(5).Text & "." & _
               vOP.ListItems.Item(A).ListSubItems(6).Text & "-" & _
               vOP.ListItems.Item(A).ListSubItems(7).Text & "." & _
               vOP.ListItems.Item(A).ListSubItems(8).Text '& "." & _
               vOP.ListItems.Item(A).ListSubItems(9).Text
        'xPro(1).Text = ""
        xPro(0).Text = A
        xPro(0).Refresh
        xPro(1).Text = NOP1
        xPro(1).Refresh
            
        vOP.ListItems.Item(A).ListSubItems(10).Text = vBangunan.ListItems.Item(A).ListSubItems(12).Text 'LUAS BUMI
        vOP.ListItems.Item(A).ListSubItems(11).Text = vBangunan.ListItems.Item(A).ListSubItems(15).Text 'KELAS BUMI
        vOP.ListItems.Item(A).ListSubItems(12).Text = vBangunan.ListItems.Item(A).ListSubItems(16).Text 'NJOP BUMI
    QJum1 = 0: QJum2 = 0: QJum3 = 0: xKet = "SISTEM"
'    vBumi.ListItems.Item(I).ListSubItems(17).Text = vBng.ListItems.Item(I).ListSubItems(64).Text
        'Menjumlahkan Nilai Bangunan yang NOP sama tetapi jumlah bangunan lebih dari 1
         For B = 1 To vBng.ListItems.Count
            If vOP.ListItems.Item(A).ListSubItems(9).Text = vBng.ListItems.Item(B).ListSubItems(64).Text Then
                'If vBng.ListItems.Item(B).ListSubItems(10).Text = "11" Then QJum1 = 0: QJum2 = 0: QJum2 = 0 ': MsgBox QJum1 & ":" & vBng.ListItems.Item(B).ListSubItems(10).Text
                QJum1 = QJum1 + (vBng.ListItems.Item(B).ListSubItems(13).Text * 1) 'Total Luas Bangunan
                If vBng.ListItems.Item(B).ListSubItems(69).Text = "INDIVIDU" Then
                    QJum2 = QJum2 + (vBng.ListItems.Item(B).ListSubItems(80).Text * 1) 'NILAI SISTEM BARU
                Else
                     QJum2 = QJum2 + (vBng.ListItems.Item(B).ListSubItems(63).Text * 1) 'NILAI SISTEM BARU
                End If
                QJum3 = QJum3 + (vBng.ListItems.Item(B).ListSubItems(24).Text * 1) 'Total NJOP Bangunan
                xKet = vBng.ListItems.Item(B).ListSubItems(69).Text
                'vBng.ListItems.Item(B).ListSubItems(10).Text
                vOP.ListItems.Item(A).ListSubItems(19).Text = vBng.ListItems.Item(B).ListSubItems(10).Text
            End If
            
            
        Next
        vOP.ListItems.Item(A).ListSubItems(13).Text = QJum1
        'vOP.ListItems.Item(A).ListSubItems(13).Text = QJum3
        vOP.ListItems.Item(A).ListSubItems(14).Text = Format(QJum2, "#,#0.00")
        vOP.ListItems.Item(A).ListSubItems(18).Text = xKet
Next
rCount = 0
For H = 1 To vOP.ListItems.Count
    If vOP.ListItems.Item(H).ListSubItems(13).Text = "" Or vOP.ListItems.Item(H).ListSubItems(13).Text = 0 Then
        rCount = rCount + 1
    End If
Next
QSTR = "SELECT * FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG ='" & xTB & "'"
openDB (QSTR)
        If vBangunan.ListItems.Count > 1 Then
            pNilai.Max = vBangunan.ListItems.Count
            pNilai.Min = 1
        End If
tJum.Refresh
tJum.Text = rCount
W = 0
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
LNilai.Visible = False
Do While Not rPajak.EOF
    W = W + 1
    'If A = pNilai.Max Then A = 1
        LNilai.Refresh
    For J = 1 To vOP.ListItems.Count
        LNilai.Visible = True
        LNilai.Caption = "7/13 - Proses Klasifikasi Bangunan : " & Round(W / pNilai.Max * 100, 0) & "%" '" & Titik3(j)
        pNilai.Value = W
        LNilai.Visible = False
    If vOP.ListItems.Item(J).ListSubItems(14).Text * 1 <> 0 Or vOP.ListItems.Item(J).ListSubItems(13).Text <> 0 Then
        If (vOP.ListItems.Item(J).ListSubItems(14).Text * 1 / vOP.ListItems.Item(J).ListSubItems(13).Text) >= rPajak!NILAI_MIN_BNG And (vOP.ListItems.Item(J).ListSubItems(14).Text * 1 / vOP.ListItems.Item(J).ListSubItems(13).Text) <= rPajak!NILAI_MAX_BNG Then
            vOP.ListItems.Item(J).ListSubItems(15).Text = Format(rPajak!KD_KLS_BNG, "000")
            vOP.ListItems.Item(J).ListSubItems(16).Text = Format(rPajak!NILAI_PER_M2_BNG, "#,#0")
            'If vOP.ListItems.Item(J).ListSubItems(18).Text = "INDIVIDU" Then
            '    vOP.ListItems.Item(J).ListSubItems(17).Text = Format(vOP.ListItems.Item(J).ListSubItems(14).Text, "#,#0.00")
            'Else
                vOP.ListItems.Item(J).ListSubItems(17).Text = Format(rPajak!NILAI_PER_M2_BNG * vOP.ListItems.Item(J).ListSubItems(13).Text * 1000, "#,#0")
            'End If
        End If
    End If
'        If vOP.ListItems.Item(j).ListSubItems(18).Text = "INDIVIDU" Then
'            vOP.ListItems.Item(j).ListSubItems(17).Text = Format(vOP.ListItems.Item(j).ListSubItems(14).Text, "#,#0")
'        End If
        'If vOP.ListItems.Item(J).ListSubItems(19).Text = "11" Then
            'vOP.ListItems.Item(J).ListSubItems(17).Text = 0
        'End If
        If vOP.ListItems.Item(J).ListSubItems(14).Text <= 0 Then
            vOP.ListItems.Item(J).ListSubItems(17).Text = 0
        End If
        'menghitung PBB Terhutang
        vOP.ListItems.Item(J).ListSubItems(31).Text = (vOP.ListItems.Item(J).ListSubItems(12).Text * 1) + (vOP.ListItems.Item(J).ListSubItems(17).Text * 1)
        If vOP.ListItems.Item(J).ListSubItems(31).Text * 1 >= ccMin(1) And vOP.ListItems.Item(J).ListSubItems(31).Text * 1 <= ccMax(1) Then
            vOP.ListItems.Item(J).ListSubItems(32).Text = ccTarif(1)
            vOP.ListItems.Item(J).ListSubItems(33).Text = ccTKP(1)
        Else
            vOP.ListItems.Item(J).ListSubItems(32).Text = ccTarif(2)
            vOP.ListItems.Item(J).ListSubItems(33).Text = ccTKP(2)
        End If
        If vOP.ListItems.Item(J).ListSubItems(20).Text * 1 <> 1 Or vOP.ListItems.Item(J).ListSubItems(17).Text * 1 = 0 Then vOP.ListItems.Item(J).ListSubItems(33).Text = 0
        vOP.ListItems.Item(J).ListSubItems(34).Text = PBBMin
        vOP.ListItems.Item(J).ListSubItems(35).Text = ((vOP.ListItems.Item(J).ListSubItems(31).Text * 1) - (vOP.ListItems.Item(J).ListSubItems(33).Text * 1)) * (vOP.ListItems.Item(J).ListSubItems(32).Text / 100)
        If vOP.ListItems.Item(J).ListSubItems(35).Text * 1 < 0 Then vOP.ListItems.Item(J).ListSubItems(35).Text = 0
    Next

    'frmBar.Bar1.Value = frmBar.Bar1.Value + 1
rPajak.MoveNext


Loop

For Y = 20 To 36
    vOP.ColumnHeaders(Y).Width = 1400
Next
If ccProses = 1 Then
    Q_SPPT = "SELECT * FROM QOBJEKPAJAK ORDER BY NOPQ ASC"
ElseIf ccProses = 2 Then
    Q_SPPT = "SELECT * FROM QOBJEKPAJAK WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' ORDER BY NOPQ ASC"
Else
    Q_SPPT = "SELECT * FROM QOBJEKPAJAK WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' ORDER BY NOPQ ASC"
End If
openDB (Q_SPPT)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        If rPajak.RecordCount > 1 Then
            pNilai.Max = rPajak.RecordCount
            pNilai.Min = 1
        End If

Q = 0
tJum.Refresh
tJum.Text = rPajak.RecordCount
Do While Not rPajak.EOF
Q = Q + 1
If Q > vOP.ListItems.Count Then Q = Q - 1
LNilai.Visible = True
        LNilai.Visible = True
        LNilai.Caption = "8/13 - Proses Identitas Wajib Pajak: " & Round(Q / pNilai.Max * 100, 0) & "%" '" & Titik3(j)
        LNilai.Refresh
        LNilai.Visible = False
        pNilai.Value = Q
        
For K = 1 To vOP.ListItems.Count

    
    If rPajak!NOPQ = vOP.ListItems.Item(K).ListSubItems(9).Text Then
        If IsNull(rPajak!Nm_wp) = True Then
            vOP.ListItems.Item(K).ListSubItems(21).Text = ""
        Else
            vOP.ListItems.Item(K).ListSubItems(21).Text = rPajak!Nm_wp
        End If
        If IsNull(rPajak!JALAN_WP) = True Then vOP.ListItems.Item(K).ListSubItems(22).Text = "-" Else vOP.ListItems.Item(K).ListSubItems(22).Text = rPajak!JALAN_WP
        If IsNull(rPajak!BLOK_KAV_NO_WP) = True Then vOP.ListItems.Item(K).ListSubItems(23).Text = "00" Else vOP.ListItems.Item(K).ListSubItems(23).Text = rPajak!BLOK_KAV_NO_WP
        If IsNull(rPajak!RW_WP) = True Then vOP.ListItems.Item(K).ListSubItems(24).Text = "00" Else vOP.ListItems.Item(K).ListSubItems(24).Text = rPajak!RW_WP
        If IsNull(rPajak!RT_WP) = True Then vOP.ListItems.Item(K).ListSubItems(25).Text = "00" Else vOP.ListItems.Item(K).ListSubItems(25).Text = rPajak!RT_WP
        If IsNull(rPajak!KELURAHAN_WP) = True Then vOP.ListItems.Item(K).ListSubItems(26).Text = "-" Else vOP.ListItems.Item(K).ListSubItems(26).Text = rPajak!KELURAHAN_WP
        If IsNull(rPajak!KOTA_WP) = True Then vOP.ListItems.Item(K).ListSubItems(27).Text = "-" Else vOP.ListItems.Item(K).ListSubItems(27).Text = rPajak!KOTA_WP
        If IsNull(rPajak!KD_POS_WP) = True Then vOP.ListItems.Item(K).ListSubItems(28).Text = "00000" Else vOP.ListItems.Item(K).ListSubItems(28).Text = rPajak!KD_POS_WP
        If IsNull(rPajak!NPWP) = True Then vOP.ListItems.Item(K).ListSubItems(29).Text = "-" Else vOP.ListItems.Item(K).ListSubItems(29).Text = rPajak!NPWP
        If IsNull(rPajak!NO_PERSIL) = True Then vOP.ListItems.Item(K).ListSubItems(30).Text = "00" Else vOP.ListItems.Item(K).ListSubItems(30).Text = rPajak!NO_PERSIL
        If IsNull(rPajak!SUBJEK_PAJAK_ID) = True Or (rPajak!SUBJEK_PAJAK_ID) = "" Then
            vOP.ListItems.Item(K).ListSubItems(36).Text = ""
        Else
            vOP.ListItems.Item(K).ListSubItems(36).Text = rPajak!SUBJEK_PAJAK_ID
        End If
        '----------------------------------
        vOP.ListItems.Item(K).ListSubItems(37).Text = rPajak!NO_FORMULIR_SPOP
        vOP.ListItems.Item(K).ListSubItems(38).Text = rPajak!KD_STATUS_WP
        vOP.ListItems.Item(K).ListSubItems(39).Text = rPajak!JNS_TRANSAKSI_OP 'Jenis Transaksi
        If IsNull(rPajak!NIP_PENDATA) = True Or rPajak!NIP_PENDATA = "" Then rPajak!NIP_PENDATA = "-"
        If IsNull(rPajak!NIP_PEMERIKSA_OP) = True Or rPajak!NIP_PEMERIKSA_OP = "" Then rPajak!NIP_PEMERIKSA_OP = "-"
        If IsNull(rPajak!NIP_PEREKAM_OP) = True Or rPajak!NIP_PEREKAM_OP = "" Then rPajak!NIP_PEREKAM_OP = "-"

        vOP.ListItems.Item(K).ListSubItems(40).Text = rPajak![TGL_PENDATAAN_OP] 'Tanggal Pendataan
        vOP.ListItems.Item(K).ListSubItems(41).Text = rPajak![NIP_PENDATA]  'NIP Petugas Pendata
        vOP.ListItems.Item(K).ListSubItems(42).Text = rPajak![TGL_PEMERIKSAAN_OP]  'Tanggal Pemeriksaan
        vOP.ListItems.Item(K).ListSubItems(43).Text = rPajak![NIP_PEMERIKSA_OP]  'NIP Petugas Pemeriksa
        vOP.ListItems.Item(K).ListSubItems(44).Text = rPajak![TGL_PEREKAMAN_OP]   'Tanggal Perekaman
        vOP.ListItems.Item(K).ListSubItems(45).Text = rPajak![NIP_PEREKAM_OP]  'NIP Petugas Perekam
        
        
        If IsNull(rPajak!JALAN_OP) = True Then vOP.ListItems.Item(K).ListSubItems(46).Text = "-" Else vOP.ListItems.Item(K).ListSubItems(46).Text = rPajak!JALAN_OP
        If IsNull(rPajak!BLOK_KAV_NO_OP) = True Then vOP.ListItems.Item(K).ListSubItems(47).Text = "00" Else vOP.ListItems.Item(K).ListSubItems(47).Text = rPajak!BLOK_KAV_NO_OP
        If IsNull(rPajak!RW_OP) = True Then vOP.ListItems.Item(K).ListSubItems(48).Text = "00" Else vOP.ListItems.Item(K).ListSubItems(48).Text = rPajak!RW_OP
        If IsNull(rPajak!RT_OP) = True Then vOP.ListItems.Item(K).ListSubItems(49).Text = "00" Else vOP.ListItems.Item(K).ListSubItems(49).Text = rPajak!RT_OP
        
        
'        vOP.ListItems.Item(K).ListSubItems(31).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(32).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(33).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(34).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(35).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(36).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(37).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(38).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(39).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(40).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(41).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(42).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(43).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(44).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(45).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(46).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(47).Text = rPajak!BLOK_KAV_NO_WP
'        vOP.ListItems.Item(K).ListSubItems(48).Text = rPajak!BLOK_KAV_NO_WP
    End If
    xPro(0).Text = Q
        xPro(1).Text = vOP.ListItems.Item(Q).ListSubItems(9).Text
       
Next
        xPro(0).Refresh
         xPro(1).Refresh
        
rPajak.MoveNext
Loop
call_judul_op
LNilai.Visible = True
LNilai.Caption = "Sukses!"
'SALAH:
'If Err.Number = 0 Then Exit Sub
'MsgBox Err.Number & ": " & Err.Description




'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'Salah:
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'MsgBox Err.Number & ": " & Err.Description

End Sub


Private Sub vOP_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vOP.SortKey = ColumnHeader.Index - 1
vOP.Sorted = True
vOP.Sorted = False
vOP.SortOrder = lvwAscending
End Sub

Sub sv_bumi()
On Error GoTo Salah
Dim xxKec, xxKel, xxBlok, xxUrut, xxJenis, xxJBumi, xxNo, xxZNT
Dim xxNIR, xxLuas, xxNilai
If ccProses = 1 Then
    m_SQL = "Select * From DAT_OP_BUMI "
ElseIf ccProses = 2 Then
    m_SQL = "Select * From DAT_OP_BUMI WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' "
Else
    m_SQL = "Select * From DAT_OP_BUMI WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'"
End If
openDB (m_SQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
tJum.Refresh
tJum.Text = rPajak.RecordCount
        If vBangunan.ListItems.Count > 1 Then
            pNilai.Max = vBangunan.ListItems.Count
            pNilai.Min = 1
        End If

Do While Not rPajak.EOF
   
    i = i + 1
    If i > vBangunan.ListItems.Count Then i = i - 1
    'for i=1 to vop
        If J > 4 Then J = 1
        LNilai.Visible = True
        LNilai.Caption = "9/13 - Proses Objek Bumi: " & Round(i / pNilai.Max * 100, 0) & "%" '" & Titik6(j)
        pNilai.Value = i
        LNilai.Refresh
        LNilai.Visible = False
    xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
'    xxKec = vBangunan.ListItems.Item(i).ListSubItems(4).Text
'    xxKel = vBangunan.ListItems.Item(i).ListSubItems(5).Text
'    xxBlok = vBangunan.ListItems.Item(i).ListSubItems(6).Text
'    xxUrut = vBangunan.ListItems.Item(i).ListSubItems(7).Text
'    xxJenis = vBangunan.ListItems.Item(i).ListSubItems(8).Text
'    xxNo = vBangunan.ListItems.Item(i).ListSubItems(9).Text
'    xxZNT = vBangunan.ListItems.Item(i).ListSubItems(10).Text
'    xxNIR = vBangunan.ListItems.Item(i).ListSubItems(11).Text
'    xxLuas = vBangunan.ListItems.Item(i).ListSubItems(12).Text
'    xxJBumi = vBangunan.ListItems.Item(i).ListSubItems(13).Text
'    xxNilai = vBangunan.ListItems.Item(i).ListSubItems(14).Text
'    xNOP = "12.12." & xxKec & "." & xxKel & "." & xxBlok & "-" & xxUrut & "." & xxJenis
    'iSQL = "INSERT INTO DAT_OP_BUMI(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,NO_BUMI,KD_ZNT,LUAS_BUMI,JNS_BUMI, NILAI_SISTEM_BUMI)" & _
    "Values('12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '1', '" & xxNo & "', '" & xxZNT & "', '" & xxLuas & "', '" & xxJBumi & "', '" & xxNilai & "')"
    
'rPajak.AddNew
'    rPajak![KD_PROPINSI] = vBangunan.ListItems.Item(i).ListSubItems(2).Text
'    rPajak![KD_DATI2] = vBangunan.ListItems.Item(i).ListSubItems(3).Text
'    rPajak![KD_KECAMATAN] = vBangunan.ListItems.Item(i).ListSubItems(4).Text
'    rPajak![KD_KELURAHAN] = vBangunan.ListItems.Item(i).ListSubItems(5).Text
'    rPajak![KD_BLOK] = vBangunan.ListItems.Item(i).ListSubItems(6).Text
'    rPajak![NO_URUT] = vBangunan.ListItems.Item(i).ListSubItems(7).Text
'    rPajak![KD_JNS_OP] = vBangunan.ListItems.Item(i).ListSubItems(8).Text
If rPajak![NO_BUMI] = vBangunan.ListItems.Item(i).ListSubItems(9).Text And xNOP = vBangunan.ListItems.Item(i).ListSubItems(17).Text Then
    rPajak![KD_ZNT] = vBangunan.ListItems.Item(i).ListSubItems(10).Text
'    rPajak![NIR] = vBangunan.ListItems.Item(i).ListSubItems(11).Text
    rPajak![LUAS_BUMI] = vBangunan.ListItems.Item(i).ListSubItems(12).Text
    rPajak![JNS_BUMI] = vBangunan.ListItems.Item(i).ListSubItems(13).Text
    rPajak![NILAI_SISTEM_BUMI] = Format(vBangunan.ListItems.Item(i).ListSubItems(14).Text, "#,#0")
    rPajak.Update

xPro(0).Text = i
xPro(0).Refresh
xPro(1).Text = xNOP
xPro(1).Refresh
'Next
End If
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
    
    
End Sub
Sub sv_bangunan()
On Error Resume Next
'On Error GoTo Salah
Dim xNOP
If ccProses = 1 Then
    B_SQL = "Select * From DAT_OP_BANGUNAN"
ElseIf ccProses = 2 Then
    B_SQL = "Select * From DAT_OP_BANGUNAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' "
Else
    B_SQL = "Select * From DAT_OP_BANGUNAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'"
End If
openDB (B_SQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
tJum.Refresh
tJum.Text = rPajak.RecordCount
        If rPajak.RecordCount > 1 Then
            pNilai.Max = rPajak.RecordCount
            pNilai.Min = 1
        End If

Do While Not rPajak.EOF
i = i + 1
        LNilai.Visible = True
        LNilai.Caption = "10/13 - Proses Objek Bangunan: " & Round(i / pNilai.Max * 100, 0) & "%" '" & titik5(j)
        pNilai.Value = i
        LNilai.Refresh
        LNilai.Visible = False
'rPajak.AddNew
xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
'    rPajak![KD_PROPINSI] = vBng.ListItems.Item(i).ListSubItems(2).Text
'    rPajak![KD_DATI2] = vBng.ListItems.Item(i).ListSubItems(3).Text
'    rPajak![KD_KECAMATAN] = vBng.ListItems.Item(i).ListSubItems(4).Text
'    rPajak![KD_KELURAHAN] = vBng.ListItems.Item(i).ListSubItems(5).Text
'    rPajak![KD_BLOK] = vBng.ListItems.Item(i).ListSubItems(6).Text
'    rPajak![NO_URUT] = vBng.ListItems.Item(i).ListSubItems(7).Text
'    rPajak![KD_JNS_OP] = vBng.ListItems.Item(i).ListSubItems(8).Text
    If vBng.ListItems.Item(i).ListSubItems(64).Text = xNOP And rPajak![NO_BNG] = vBng.ListItems.Item(i).ListSubItems(9).Text And rPajak![KD_JPB] = vBng.ListItems.Item(i).ListSubItems(10).Text Then
    'MsgBox vBng.ColumnHeaders(62).Text & ": " & vBng.ListItems.Item(i).ListSubItems(62).Text
    'B_SQL = "UPDATE DAT_OP_BUMI SET NILAI_SISTEM_BNG='" & vBng.ListItems.Item(i).ListSubItems(62).Text & "' WHERE (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(vBng.ListItems.Item(i).ListSubItems(64).Text) & "' And NO_BNG = '" & vBng.ListItems.Item(i).ListSubItems(9).Text & "' And kd_jpb = '" & vBng.ListItems.Item(i).ListSubItems(10).Text & "')"
    'openDB (B_SQL)
    'iSQL4 = "UPDATE DAT_OBJEK_PAJAK SET TOTAL_LUAS_BNG = '" & t_Luas & "', NJOP_BNG = '" & t_NJOP & "' where (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "')"
'    rPajak![NO_FORMULIR_LSPOP] = vBng.ListItems.Item(i).ListSubItems(72).Text '72
        'rPajak![THN_DIBANGUN_BNG] = vBng.ListItems.Item(i).ListSubItems(11).Text
        'rPajak![thn_renovasi_bng] = vBng.ListItems.Item(i).ListSubItems(12).Text
        'rPajak![LUAS_BNG] = vBng.ListItems.Item(i).ListSubItems(13).Text
        'rPajak![JML_LANTAI_BNG] = vBng.ListItems.Item(i).ListSubItems(14).Text
        'rPajak![KONDISI_BNG] = vBng.ListItems.Item(i).ListSubItems(15).Text
        'rPajak![JNS_KONSTRUKSI_BNG] = vBng.ListItems.Item(i).ListSubItems(16).Text
        'rPajak![JNS_ATAP_BNG] = vBng.ListItems.Item(i).ListSubItems(17).Text
        'rPajak![KD_DINDING] = vBng.ListItems.Item(i).ListSubItems(18).Text
        'rPajak![KD_LANTAI] = vBng.ListItems.Item(i).ListSubItems(19).Text
        'rPajak![KD_LANGIT_LANGIT] = vBng.ListItems.Item(i).ListSubItems(20).Text
        rPajak![K_UTAMA] = vBng.ListItems.Item(i).ListSubItems(81).Text
        rPajak![K_MATERIAL] = vBng.ListItems.Item(i).ListSubItems(82).Text
        rPajak![K_FASILITAS] = vBng.ListItems.Item(i).ListSubItems(83).Text
        rPajak![J_SUSUT] = vBng.ListItems.Item(i).ListSubItems(84).Text
        rPajak![K_SUSUT] = vBng.ListItems.Item(i).ListSubItems(85).Text
        rPajak![K_NON_SUSUT] = vBng.ListItems.Item(i).ListSubItems(86).Text
        rPajak![NILAI_SISTEM_BNG] = Format(vBng.ListItems.Item(i).ListSubItems(63).Text, "#,#0")
'    rPajak![JNS_TRANSAKSI_BNG] = vBng.ListItems.Item(i).ListSubItems(73).Text '73
'    rPajak![TGL_PENDATAAN_BNG] = vBng.ListItems.Item(i).ListSubItems(74).Text '74
'    rPajak![NIP_PENDATA_BNG] = vBng.ListItems.Item(i).ListSubItems(75).Text '75
'    rPajak![TGL_PEMERIKSAAN_BNG] = vBng.ListItems.Item(i).ListSubItems(76).Text '76
'    rPajak![NIP_PEMERIKSA_BNG] = vBng.ListItems.Item(i).ListSubItems(77).Text '77
'    rPajak![TGL_PEREKAMAN_BNG] = vBng.ListItems.Item(i).ListSubItems(78).Text '78
'    rPajak![NIP_PEREKAM_BNG] = vBng.ListItems.Item(i).ListSubItems(79).Text '79
        rPajak.Update
        
        xPro(0).Text = i
        xPro(0).Refresh
        xPro(1).Text = xNOP
        xPro(1).Refresh
End If
'Next
rPajak.MoveNext
Loop
'SALAH:
'If Err.Number = 0 Then Exit Sub
'MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error"
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'Salah:
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'MsgBox Err.Number & ": " & Err.Description
'
End Sub
Sub sv_Objek()
On Error GoTo Salah
'On Error Resume Next
Dim xxProp, xxKab, xxKec, xxKel, xxBlok, xxUrut, xxJenis, xNOP1, xxKelas1, xxKelas2
Dim xxJatuh
Dim LBumi, LBng, NBumi, NBng, t_NJOP, xNJOPTKP, xNJKP, xHutang, xKurang, xBayar, xxTarif, xxPBBMin
Dim xNOP
'For i = 1 To vOP.ListItems.Count
'    'Cek Nomor Identitas Subjek Pajak di Database DAT_OBJEK_PAJAK
'    If vOP.ListItems.Item(i).ListSubItems(36).Text = "" Or vOP.ListItems.Item(i).ListSubItems(36).Text = "0" Or IsNull(vOP.ListItems.Item(i).ListSubItems(36).Text) = True Then
'        'MsgBox "Cek Kembali Pendataan Objek Bumi dan Bangunan" & _
'    vbCrLf & "Proses tidak selesai dengan sempurna...!", vbCritical, "Error"
'        'MsgBox vOP.ListItems.Item(i).ListSubItems(36).Text & "//" & vOP.ListItems.Item(i).ListSubItems(36).Text & "//" & IsNull(vOP.ListItems.Item(i).ListSubItems(36).Text) = True
'       LNilai.Visible = True
'       LNilai.Caption = "Proses Gagal!!!, Silahkan Diulang!"
'       LNilai.Refresh
'       TANYA = MsgBox("Apa anda ingin mengembalikan Objek Pajak terakhir??", vbInformation + vbYesNo, "Error")
'       If TANYA = vbNo Then
'           Exit Sub
'       Else
'
'           O_SQL1 = "DELETE FROM DAT_OBJEK_PAJAK"
'           openDB (O_SQL1)
'           O_SQL2 = "insert into DAT_OBJEK_PAJAK select * from backup_Nilai_OP"
'           openDB (O_SQL2)
'           tanya2 = MsgBox("Apa anda melanjutkan proses penilaian. ??", vbInformation + vbYesNo, "Error")
'           If tanya2 = vbNo Then
'               Exit Sub
'           Else
'               GoTo lanjut
'           End If
'       End If
'    End If
'Next
'lanjut:

O_SQL1 = "B_OP '" & ccProses & "','" & Left(Trim(ccKec.Text), 3) & "','" & Left(Trim(ccKel.Text), 3) & "'"
openDB (O_SQL1)
If ccProses = "1" Then
    O_SQl = "Select * From DAT_OBJEK_PAJAK "
    
'    O_SQL1 = "DELETE FROM DAT_OBJEK_PAJAK"
  '  O_SQL1 = "B_OP '1'"
ElseIf ccProses = "2" Then
    O_SQl = "Select * From DAT_OBJEK_PAJAK WHERE KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "'"
    'O_SQL1 = "DELETE FROM DAT_OBJEK_PAJAK WHERE KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "'"
 '   O_SQL1 = "B_OP '2','" & Left(Trim(ccKec.Text), 3) & "'"
Else
    O_SQl = "Select * From DAT_OBJEK_PAJAK WHERE KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(ccKel.Text), 3) & "'"
    'O_SQL1 = "DELETE From DAT_OBJEK_PAJAK WHERE KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(ccKel.Text), 3) & "'"
'    O_SQL1 = "B_OP '3','" & Left(Trim(ccKec.Text), 3) & "','" & Left(Trim(ccKel.Text), 3) & "'"
End If

openDB (O_SQl)

If rPajak.RecordCount > 0 Then rPajak.MoveFirst
J = 0: i = 0

        If vOP.ListItems.Count > 1 Then
            pNilai.Max = vOP.ListItems.Count
            pNilai.Min = 1
        End If
tJum.Refresh
tJum.Text = vOP.ListItems.Count
'Do While Not rPajak.EOF
    For i = 1 To vOP.ListItems.Count
    LNilai.Visible = True
        LNilai.Caption = "12/13 - Simpan Data Objek Pajak: " & Round(i / pNilai.Max * 100, 0) & "%" '" & TITIK4(j)
        pNilai.Value = i
        LNilai.Refresh
        LNilai.Visible = False
    rPajak.AddNew
    rPajak![KD_PROPINSI] = vOP.ListItems.Item(i).ListSubItems(2).Text
    rPajak![KD_DATI2] = vOP.ListItems.Item(i).ListSubItems(3).Text
    rPajak![KD_KECAMATAN] = vOP.ListItems.Item(i).ListSubItems(4).Text
    rPajak![KD_KELURAHAN] = vOP.ListItems.Item(i).ListSubItems(5).Text
    rPajak![KD_BLOK] = vOP.ListItems.Item(i).ListSubItems(6).Text
    rPajak![NO_URUT] = vOP.ListItems.Item(i).ListSubItems(7).Text
    rPajak![KD_JNS_OP] = vOP.ListItems.Item(i).ListSubItems(8).Text
    '--------------------------------------------------------------------------
    rPajak![SUBJEK_PAJAK_ID] = vOP.ListItems.Item(i).ListSubItems(36).Text
    rPajak![NO_FORMULIR_SPOP] = vOP.ListItems.Item(i).ListSubItems(37).Text
    rPajak![NO_PERSIL] = vOP.ListItems.Item(i).ListSubItems(30).Text
    rPajak![JALAN_OP] = vOP.ListItems.Item(i).ListSubItems(46).Text
    rPajak![BLOK_KAV_NO_OP] = vOP.ListItems.Item(i).ListSubItems(47).Text
    rPajak![RW_OP] = vOP.ListItems.Item(i).ListSubItems(48).Text
    rPajak![RT_OP] = vOP.ListItems.Item(i).ListSubItems(49).Text
    rPajak![KD_STATUS_CABANG] = 0 'vOP.ListItems.Item(i).ListSubItems(14).Text
    rPajak![KD_STATUS_WP] = vOP.ListItems.Item(i).ListSubItems(38).Text '
    '-----------------------------------------------------------------------
    rPajak![TOTAL_LUAS_BUMI] = vOP.ListItems.Item(i).ListSubItems(10).Text
    rPajak![TOTAL_LUAS_BNG] = vOP.ListItems.Item(i).ListSubItems(13).Text
    rPajak![NJOP_BUMI] = Format(vOP.ListItems.Item(i).ListSubItems(12).Text * 1, "#,#0")
    rPajak![NJOP_BNG] = Format(vOP.ListItems.Item(i).ListSubItems(17).Text * 1, "#,#0")
    
    rPajak![STATUS_PETA_OP] = 1 'vOP.ListItems.Item(i).ListSubItems(14).Text
    rPajak![JNS_TRANSAKSI_OP] = vOP.ListItems.Item(i).ListSubItems(39).Text
    rPajak![TGL_PENDATAAN_OP] = Format(vOP.ListItems.Item(i).ListSubItems(40).Text, "DD/MM/YYYY")
    rPajak![NIP_PENDATA] = vOP.ListItems.Item(i).ListSubItems(41).Text
    rPajak![TGL_PEMERIKSAAN_OP] = Format(vOP.ListItems.Item(i).ListSubItems(42).Text, "DD/MM/YYYY")
    rPajak![NIP_PEMERIKSA_OP] = vOP.ListItems.Item(i).ListSubItems(43).Text
    rPajak![TGL_PEREKAMAN_OP] = Format(vOP.ListItems.Item(i).ListSubItems(44).Text, "DD/MM/YYYY")
    rPajak![NIP_PEREKAM_OP] = vOP.ListItems.Item(i).ListSubItems(45).Text
    rPajak.Update
        
'        If ccProses = 1 Then
'        cc_SQL = "UPDATE DAT_OBJEK_PAJAK SET NJOP_BUMI='" & Format(vOP.ListItems.Item(i).ListSubItems(12).Text * 1, "#,#0") & "',NJOP_BNG='" & Format(vOP.ListItems.Item(i).ListSubItems(17).Text * 1, "#,#0") & "' " & _
'                "WHERE Trim([KD_PROPINSI]) + '.' + Trim([KD_DATI2]) + '.' + Trim([KD_KECAMATAN]) + '.' + Trim([KD_KELURAHAN]) + '.' + Trim([KD_BLOK]) + '-' + Trim([NO_URUT]) + '.' + Trim([KD_JNS_OP])='" & Trim(vOP.ListItems.Item(i).ListSubItems(9).Text) & "'"
'        ElseIf ccProses = 2 Then
'        cc_SQL = "UPDATE DAT_OBJEK_PAJAK SET NJOP_BUMI='" & Format(vOP.ListItems.Item(i).ListSubItems(12).Text * 1, "#,#0") & "',NJOP_BNG='" & Format(vOP.ListItems.Item(i).ListSubItems(17).Text * 1, "#,#0") & "' " & _
'                "WHERE KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' and Trim([KD_PROPINSI]) + '.' + Trim([KD_DATI2]) + '.' + Trim([KD_KECAMATAN]) + '.' + Trim([KD_KELURAHAN]) + '.' + Trim([KD_BLOK]) + '-' + Trim([NO_URUT]) + '.' + Trim([KD_JNS_OP])='" & Trim(vOP.ListItems.Item(i).ListSubItems(9).Text) & "'"
'        Else
'        cc_SQL = "UPDATE DAT_OBJEK_PAJAK SET NJOP_BUMI='" & Format(vOP.ListItems.Item(i).ListSubItems(12).Text * 1, "#,#0") & "',NJOP_BNG='" & Format(vOP.ListItems.Item(i).ListSubItems(17).Text * 1, "#,#0") & "' " & _
'                "WHERE KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(ccKel.Text), 3) & "' and Trim([KD_PROPINSI]) + '.' + Trim([KD_DATI2]) + '.' + Trim([KD_KECAMATAN]) + '.' + Trim([KD_KELURAHAN]) + '.' + Trim([KD_BLOK]) + '-' + Trim([NO_URUT]) + '.' + Trim([KD_JNS_OP])='" & Trim(vOP.ListItems.Item(i).ListSubItems(9).Text) & "'"
'        End If
'        openDB (cc_SQL)
'End If
'Next

    xPro(0).Text = i
    xPro(0).Refresh
    xPro(1).Text = xNOP
    xPro(1).Refresh
Next
'rPajak.MoveNext
'Loop
'Untuk Jenis Bumi=4 Tidak Diinsert ke SPPT (vbangunan.listitems.item(L).listsubitems(13).text
iSQL = "select * from SPPT WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' and THN_PAJAK_SPPT='" & ccTahun.Text & "'"
openDB (iSQL)
'Simpan Data Ke Tabel SPPT
'callTarif
'J = 0
'Do While Not rPajak.EOF
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        If vOP.ListItems.Count > 1 Then
            pNilai.Max = vOP.ListItems.Count
            pNilai.Min = 1
        End If
tJum.Refresh
tJum.Text = vOP.ListItems.Count
a1 = 0
For L = 1 To vOP.ListItems.Count
        
        LNilai.Visible = True
        LNilai.Caption = "13/13 - Proses Pembentukan SPPT: " & Round(L / pNilai.Max * 100, 0) & "%" '" & TITIK4(j)
        LNilai.Refresh
        pNilai.Value = L
        LNilai.Visible = False
    xxProp = vOP.ListItems.Item(L).ListSubItems(2).Text
    xxKab = vOP.ListItems.Item(L).ListSubItems(3).Text
    xxKec = vOP.ListItems.Item(L).ListSubItems(4).Text
    xxKel = vOP.ListItems.Item(L).ListSubItems(5).Text
    xxBlok = vOP.ListItems.Item(L).ListSubItems(6).Text
    xxUrut = vOP.ListItems.Item(L).ListSubItems(7).Text
    xxJenis = vOP.ListItems.Item(L).ListSubItems(8).Text
    xNOP1 = xxProp & "." & xxKab & "." & xxKec & "." & xxKel & "." & xxBlok & "-" & xxUrut & "." & xxJenis
    xxKelas1 = vOP.ListItems.Item(L).ListSubItems(11).Text
    xxKelas2 = vOP.ListItems.Item(L).ListSubItems(15).Text
    xxJatuh = Format(Now, "dd/mm/yyyy")
    LBumi = vOP.ListItems.Item(L).ListSubItems(10).Text
    LBng = vOP.ListItems.Item(L).ListSubItems(13).Text
    NBumi = vOP.ListItems.Item(L).ListSubItems(12).Text
    NBng = vOP.ListItems.Item(L).ListSubItems(17).Text
    t_NJOP = vOP.ListItems.Item(L).ListSubItems(31).Text '(NBumi * 1) + (NBng * 1)

'    If t_NJOP >= ccMin(1) * 1 And t_NJOP <= ccMax(1) * 1 Then
'        xxTarif = ccTarif(1) / 100
'        xNJOPTKP = ccTKP(1) * 1
'    Else
'        xxTarif = ccTarif(2) / 100
'        xNJOPTKP = ccTKP(2) * 1
'    End If
    xNJOPTKP = vOP.ListItems.Item(L).ListSubItems(33).Text
    xNJKP = 0
    'If vOP.ListItems.Item(L).ListSubItems(20).Text * 1 <> 1 Then xNJOPTK = 0
    'xHutang = ((t_NJOP * 1) - (xNJOPTKP * 1)) * (xxTarif * 1) 'PBB Terutang
    xxPBBMin = vOP.ListItems.Item(L).ListSubItems(34).Text
    xHutang = Format(vOP.ListItems.Item(L).ListSubItems(35).Text, "#,#0")
    xKurang = 0 'Faktor Pengurang
    xBayar = Format((xHutang * 1) - (xKurang * 1), "#,#0") 'PBB Yang harus dibayar
    If xBayar * 1 < xxPBBMin * 1 Then xBayar = xxPBBMin * 1
    If vOP.ListItems.Item(L).ListSubItems(20).Text <> "4" Then
'If xNOP = vOP.ListItems.Item(i).ListSubItems(9).Text Then
a1 = a1 + 1
    'iSQL = "INSERT INTO SPPT(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,THN_PAJAK_SPPT,NM_WP_SPPT,JLN_WP_SPPT,BLOK_KAV_NO_WP_SPPT,RW_WP_SPPT,RT_WP_SPPT,KELURAHAN_WP_SPPT,KOTA_WP_SPPT,KD_POS_WP_SPPT,NPWP_SPPT," & _
    "NO_PERSIL_SPPT,KD_KLS_TANAH,THN_AWAL_KLS_TANAH,KD_KLS_BNG,THN_AWAL_KLS_BNG,TGL_JATUH_TEMPO_SPPT,LUAS_BUMI_SPPT,LUAS_BNG_SPPT, NJOP_BUMI_SPPT,NJOP_BNG_SPPT,NJOP_SPPT,NJOPTKP_SPPT,NJKP_SPPT,PBB_TERHUTANG_SPPT,FAKTOR_PENGURANG_SPPT," & _
    "PBB_YG_HARUS_DIBAYAR_SPPT,STATUS_PEMBAYARAN_SPPT,STATUS_TAGIHAN_SPPT,STATUS_CETAK_SPPT,TGL_TERBIT_SPPT,TGL_CETAK_SPPT,NIP_PENCETAK_SPPT,SIKLUS_SPPT,KD_KANWIL_BANK,KD_KPPBB_BANK,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,KD_TP,PROSES)" & _
    "Values('12', '12', '" & xxKec & "', '" & xxKel & "', '" & xxBlok & "', '" & xxUrut & "', '" & xxJenis & "', '" & ccTahun.Text & "', '" & vOP.ListItems.Item(L).ListSubItems(21).Text & "', '" & vOP.ListItems.Item(L).ListSubItems(22).Text & "', '" & vOP.ListItems.Item(L).ListSubItems(23).Text & "', '" & vOP.ListItems.Item(L).ListSubItems(24).Text & "', '" & vOP.ListItems.Item(L).ListSubItems(25).Text & "','" & vOP.ListItems.Item(L).ListSubItems(26).Text & "','" & vOP.ListItems.Item(L).ListSubItems(27).Text & "','" & vOP.ListItems.Item(L).ListSubItems(28).Text & "','" & vOP.ListItems.Item(L).ListSubItems(29).Text & "','" & vOP.ListItems.Item(L).ListSubItems(30).Text & "'," & _
    " '" & xxKelas1 & " ','" & xTT & "','" & xxKelas2 & "','" & xTB & "','" & xxJatuh & "','" & LBumi & "','" & LBng & "','" & NBumi & "','" & NBng & "','" & t_NJOP & "','" & xNJOPTKP & "','" & xNJKP & "','" & xHutang & "','" & xKurang & "','" & xBayar & "','0','0','0','" & xxJatuh & "','" & xxJatuh & "','000000', " & _
    " 1,'01','16','04','01','93','N')"
    'openDB (iSQL)

    rPajak.AddNew
    rPajak!KD_PROPINSI = "12"
rPajak!KD_DATI2 = "12"
rPajak!KD_KECAMATAN = xxKec
rPajak!KD_KELURAHAN = xxKel
rPajak!KD_BLOK = xxBlok
rPajak!NO_URUT = xxUrut
rPajak!KD_JNS_OP = xxJenis
rPajak!THN_PAJAK_SPPT = ccTahun.Text
rPajak!NM_WP_SPPT = "-"
rPajak!JLN_WP_SPPT = "-"
rPajak!BLOK_KAV_NO_WP_SPPT = "00"
rPajak!RW_WP_SPPT = "00"
rPajak!RT_WP_SPPT = "00"
rPajak!KELURAHAN_WP_SPPT = "-"
rPajak!KOTA_WP_SPPT = "-"
rPajak!KD_POS_WP_SPPT = "-"
rPajak!NPWP_SPPT = "0"
rPajak!NO_PERSIL_SPPT = "00"
rPajak!KD_KLS_TANAH = xxKelas1
rPajak!THN_AWAL_KLS_TANAH = xTT
rPajak!KD_KLS_BNG = xxKelas2
rPajak!THN_AWAL_KLS_BNG = xTB
rPajak!TGL_JATUH_TEMPO_SPPT = xxJatuh
rPajak!LUAS_BUMI_SPPT = LBumi
rPajak!LUAS_BNG_SPPT = LBng
rPajak!NJOP_BUMI_SPPT = NBumi
rPajak!NJOP_BNG_SPPT = NBng
rPajak!NJOP_SPPT = t_NJOP
rPajak!NJOPTKP_SPPT = xNJOPTKP
rPajak!NJKP_SPPT = xNJKP
rPajak!PBB_TERHUTANG_SPPT = xHutang
rPajak!FAKTOR_PENGURANG_SPPT = xKurang
rPajak!PBB_YG_HARUS_DIBAYAR_SPPT = xBayar
rPajak!STATUS_PEMBAYARAN_SPPT = "0"
rPajak!STATUS_TAGIHAN_SPPT = "0"
rPajak!STATUS_CETAK_SPPT = "0"
rPajak!TGL_TERBIT_SPPT = xxJatuh
rPajak!TGL_CETAK_SPPT = xxJatuh
rPajak!NIP_PENCETAK_SPPT = "000000"
rPajak!SIKLUS_SPPT = 1
rPajak!KD_KANWIL_BANK = "01"
rPajak!KD_KPPBB_BANK = "16"
rPajak!KD_BANK_TUNGGAL = "04"
rPajak!KD_BANK_PERSEPSI = "01"
rPajak!KD_TP = "93"
rPajak!PROSES = "N"
rPajak.Update

    xPro(0).Text = a1
    xPro(0).Refresh
    xPro(1).Text = xNOP1
    xPro(1).Refresh
End If
'End If
Next
'rPajak.MoveNext
'Loop
LNilai.Visible = True
LNilai.Caption = "Sukses!"
pNilai.Visible = False
'SALAH:
'If Err.Number = 0 Then Exit Sub
'MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error"

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
Sub sv_individu()
On Error GoTo Salah
Dim xxKec, xxKel, xxBlok, xxUrut, xxJenis, xxJBumi, xxNo, xxZNT
Dim xxNIR, xxLuas, xxNilai
Dim xNOP
If ccProses = 1 Then
    I_SQL = "Select * From DAT_NILAI_INDIVIDU"
ElseIf ccProses = 2 Then
    I_SQL = "Select * From DAT_NILAI_INDIVIDU WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "'"
Else
    I_SQL = "Select * From DAT_NILAI_INDIVIDU WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'"
End If



openDB (I_SQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
tJum.Refresh
tJum.Text = rPajak.RecordCount
        If vBng.ListItems.Count > 1 Then
            pNilai.Max = vBng.ListItems.Count
            pNilai.Min = 1
        End If

Do While Not rPajak.EOF

i = i + 1
        If J > 4 Then J = 1
        LNilai.Visible = True
        LNilai.Caption = "11/13 - Penilaian Objek secara individu: " & Round(i / pNilai.Max * 100, 0) & "%" '" & Titik7(J)
        LNilai.Refresh
        LNilai.Visible = False
        pNilai.Value = i
        xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
        
'rPajak.AddNew
'    rPajak![KD_PROPINSI] = vBng.ListItems.Item(i).ListSubItems(2).Text
'    rPajak![KD_DATI2] = vBng.ListItems.Item(i).ListSubItems(3).Text
'    rPajak![KD_KECAMATAN] = vBng.ListItems.Item(i).ListSubItems(4).Text
'    rPajak![KD_KELURAHAN] = vBng.ListItems.Item(i).ListSubItems(5).Text
'    rPajak![KD_BLOK] = vBng.ListItems.Item(i).ListSubItems(6).Text
'    rPajak![NO_URUT] = vBng.ListItems.Item(i).ListSubItems(7).Text
'    rPajak![KD_JNS_OP] = vBng.ListItems.Item(i).ListSubItems(8).Text
   If vBng.ListItems.Item(i).ListSubItems(64).Text = xNOP And rPajak![NO_BNG] = vBng.ListItems.Item(i).ListSubItems(9).Text Then
'    rPajak![NO_FORMULIR_INDIVIDUAL] = vBng.ListItems.Item(i).ListSubItems(72).Text
        rPajak![NILAI_INDIVIDU] = vBng.ListItems.Item(i).ListSubItems(80).Text
        rPajak.Update
        xPro(0).Text = i
        xPro(0).Refresh
        xPro(1).Text = xNOP
        xPro(1).Refresh
   End If
'    rPajak![TGL_PENILAIAN_INDIVIDU] = vBng.ListItems.Item(i).ListSubItems(74).Text
'    rPajak![NIP_PENILAI_INDIVIDU] = vBng.ListItems.Item(i).ListSubItems(75).Text
'    rPajak![TGL_PEMERIKSAAN_INDIVIDU] = vBng.ListItems.Item(i).ListSubItems(76).Text
'    rPajak![NIP_PEMERIKSA_INDIVIDU] = vBng.ListItems.Item(i).ListSubItems(77).Text
'    rPajak![TGL_REKAM_NILAI_INDIVIDU] = vBng.ListItems.Item(i).ListSubItems(78).Text
'    rPajak![NIP_PEREKAM_INDIVIDU] = vBng.ListItems.Item(i).ListSubItems(79).Text

'xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
'Next
rPajak.MoveNext
Loop
'SALAH:
'If Err.Number = 0 Then Exit Sub
'MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Sub call_judul_op()
On Error Resume Next
vOP.ColumnHeaders(3).Text = "PROP"
                vOP.ColumnHeaders(4).Text = "KAB"
                vOP.ColumnHeaders(5).Text = "KEC"
                vOP.ColumnHeaders(6).Text = "KEL"
                vOP.ColumnHeaders(7).Text = "BLOK"
                vOP.ColumnHeaders(8).Text = "URUT"
                vOP.ColumnHeaders(9).Text = "KODE"
                vOP.ColumnHeaders(10).Text = "NOP"
                vOP.ColumnHeaders(11).Text = "LUAS_BUMI"
                vOP.ColumnHeaders(12).Text = "KLS_BUMI"
                vOP.ColumnHeaders(13).Text = "NJOP_BUMI"
                vOP.ColumnHeaders(14).Text = "LUAS_BNG"
                vOP.ColumnHeaders(15).Text = "NILAI SISTEM"
                vOP.ColumnHeaders(16).Text = "KLS_BNG"
                vOP.ColumnHeaders(17).Text = "NJOP_PER_M2"
                vOP.ColumnHeaders(18).Text = "NJOP_BNG"
                vOP.ColumnHeaders(19).Text = "KETERANGAN"
                vOP.ColumnHeaders(20).Text = "JPB"
                vOP.ColumnHeaders(21).Text = "Jenis Bumi"
                vOP.ColumnHeaders(2).Width = 800
                vOP.ColumnHeaders(3).Width = 700
                vOP.ColumnHeaders(4).Width = 700
                vOP.ColumnHeaders(5).Width = 700
                vOP.ColumnHeaders(6).Width = 700
                vOP.ColumnHeaders(7).Width = 700
                vOP.ColumnHeaders(8).Width = 700
                vOP.ColumnHeaders(9).Width = 700
                vOP.ColumnHeaders(10).Width = 1800
                vOP.ColumnHeaders(11).Width = 1000
                vOP.ColumnHeaders(12).Width = 1000
                vOP.ColumnHeaders(13).Width = 1400
                vOP.ColumnHeaders(14).Width = 1400
                vOP.ColumnHeaders(15).Width = 1400
                vOP.ColumnHeaders(16).Width = 700
                vOP.ColumnHeaders(17).Width = 1400
                vOP.ColumnHeaders(18).Width = 1400
                vOP.ColumnHeaders(19).Width = 1400
                vOP.ColumnHeaders(20).Width = 1400
                vOP.ColumnHeaders(21).Width = 1400
                vOP.ColumnHeaders(22).Width = 1400
End Sub
Sub call_judul_fas()
On Error Resume Next
vFAS.ColumnHeaders(3).Text = "PROP"
                vFAS.ColumnHeaders(4).Text = "KAB"
                vFAS.ColumnHeaders(5).Text = "KEC"
                vFAS.ColumnHeaders(6).Text = "KEL"
                vFAS.ColumnHeaders(7).Text = "BLOK"
                vFAS.ColumnHeaders(8).Text = "URUT"
                vFAS.ColumnHeaders(9).Text = "KODE"
                vFAS.ColumnHeaders(10).Text = "NO BNG"
                vFAS.ColumnHeaders(11).Text = "FAS"
                vFAS.ColumnHeaders(12).Text = "JLH_SATUAN"
                vFAS.ColumnHeaders(13).Text = "DBKB"
                vFAS.ColumnHeaders(14).Text = "JUMLAH"
                vFAS.ColumnHeaders(15).Text = "04"
                vFAS.ColumnHeaders(16).Text = "05"
                vFAS.ColumnHeaders(17).Text = "06"
                vFAS.ColumnHeaders(18).Text = "07"
                vFAS.ColumnHeaders(19).Text = "08"
                vFAS.ColumnHeaders(20).Text = "DBKB Fas"
                vFAS.ColumnHeaders(21).Text = "10"
                vFAS.ColumnHeaders(22).Text = "11"
                vFAS.ColumnHeaders(23).Text = "12"
                vFAS.ColumnHeaders(24).Text = "13"
                vFAS.ColumnHeaders(25).Text = "14"
                vFAS.ColumnHeaders(26).Text = "15"
                vFAS.ColumnHeaders(27).Text = "16"
                vFAS.ColumnHeaders(28).Text = "17"
                vFAS.ColumnHeaders(29).Text = "18"
                vFAS.ColumnHeaders(30).Text = "19"
                vFAS.ColumnHeaders(31).Text = "20"
                vFAS.ColumnHeaders(32).Text = "21"
                vFAS.ColumnHeaders(33).Text = "22"
                vFAS.ColumnHeaders(34).Text = "23"
                vFAS.ColumnHeaders(35).Text = "24"
                vFAS.ColumnHeaders(36).Text = "25"
                vFAS.ColumnHeaders(37).Text = "26"
                vFAS.ColumnHeaders(38).Text = "27"
                vFAS.ColumnHeaders(39).Text = "28"
                vFAS.ColumnHeaders(40).Text = "29"
                vFAS.ColumnHeaders(41).Text = "30"
                vFAS.ColumnHeaders(42).Text = "31"
                vFAS.ColumnHeaders(43).Text = "32"
                vFAS.ColumnHeaders(44).Text = "33"
                vFAS.ColumnHeaders(45).Text = "34"
                vFAS.ColumnHeaders(46).Text = "35"
                vFAS.ColumnHeaders(47).Text = "36"
                vFAS.ColumnHeaders(48).Text = "37"
                vFAS.ColumnHeaders(49).Text = "38"
                vFAS.ColumnHeaders(50).Text = "39"
                vFAS.ColumnHeaders(51).Text = "40"
                vFAS.ColumnHeaders(52).Text = "41"
                vFAS.ColumnHeaders(53).Text = "42"
                vFAS.ColumnHeaders(54).Text = "43"
                vFAS.ColumnHeaders(55).Text = "44"
                vFAS.ColumnHeaders(56).Text = "45"
                vFAS.ColumnHeaders(57).Text = "DBKB"
                
                vFAS.ColumnHeaders(2).Width = 800
                vFAS.ColumnHeaders(3).Width = 700
                vFAS.ColumnHeaders(4).Width = 700
                vFAS.ColumnHeaders(5).Width = 700
                vFAS.ColumnHeaders(6).Width = 700
                vFAS.ColumnHeaders(7).Width = 700
                vFAS.ColumnHeaders(8).Width = 700
                vFAS.ColumnHeaders(9).Width = 700
                vFAS.ColumnHeaders(10).Width = 700
                vFAS.ColumnHeaders(11).Width = 700
                vFAS.ColumnHeaders(12).Width = 1200
                vFAS.ColumnHeaders(13).Width = 1000
                vFAS.ColumnHeaders(14).Width = 1200
                vFAS.ColumnHeaders(15).Width = 1400
                vFAS.ColumnHeaders(16).Width = 700
                vFAS.ColumnHeaders(17).Width = 700
                vFAS.ColumnHeaders(18).Width = 700
                vFAS.ColumnHeaders(19).Width = 700
                vFAS.ColumnHeaders(20).Width = 700
                vFAS.ColumnHeaders(21).Width = 700
                vFAS.ColumnHeaders(22).Width = 1900
                vFAS.ColumnHeaders(23).Width = 700
                vFAS.ColumnHeaders(24).Width = 1400
                vFAS.ColumnHeaders(25).Width = 1900
                vFAS.ColumnHeaders(26).Width = 1900
                vFAS.ColumnHeaders(3).Alignment = lvwColumnCenter: vFAS.ColumnHeaders(4).Alignment = lvwColumnCenter
                vFAS.ColumnHeaders(5).Alignment = lvwColumnCenter: vFAS.ColumnHeaders(6).Alignment = lvwColumnCenter
                vFAS.ColumnHeaders(7).Alignment = lvwColumnCenter: vFAS.ColumnHeaders(8).Alignment = lvwColumnCenter
                vFAS.ColumnHeaders(9).Alignment = lvwColumnCenter: vFAS.ColumnHeaders(10).Alignment = lvwColumnCenter
                vFAS.ColumnHeaders(11).Alignment = lvwColumnCenter: vFAS.ColumnHeaders(12).Alignment = lvwColumnRight
                vFAS.ColumnHeaders(13).Alignment = lvwColumnRight: vFAS.ColumnHeaders(14).Alignment = lvwColumnRight
                vFAS.ColumnHeaders(15).Alignment = lvwColumnRight: vFAS.ColumnHeaders(16).Alignment = lvwColumnCenter
                vFAS.ColumnHeaders(17).Alignment = lvwColumnRight
End Sub
Sub judul_nilai_bng()
On Error Resume Next
vBng.ColumnHeaders(3).Text = "PROP"
                vBng.ColumnHeaders(4).Text = "KAB"
                vBng.ColumnHeaders(5).Text = "KEC"
                vBng.ColumnHeaders(6).Text = "KEL"
                vBng.ColumnHeaders(7).Text = "BLOK"
                vBng.ColumnHeaders(8).Text = "URUT"
                vBng.ColumnHeaders(9).Text = "KODE"
                vBng.ColumnHeaders(10).Text = "NO BNG"
                vBng.ColumnHeaders(11).Text = "JPB"
                vBng.ColumnHeaders(12).Text = "T_BANGUN"
                vBng.ColumnHeaders(13).Text = "T_RENOVASI"
                vBng.ColumnHeaders(14).Text = "LUAS"
                vBng.ColumnHeaders(15).Text = "LANTAI"
                vBng.ColumnHeaders(16).Text = "KONDISI"
                vBng.ColumnHeaders(17).Text = "KONSTRUKSI"
                vBng.ColumnHeaders(18).Text = "ATAP"
                vBng.ColumnHeaders(19).Text = "DINDING"
                vBng.ColumnHeaders(20).Text = "LANTAI"
                vBng.ColumnHeaders(21).Text = "LANGIT2"
                vBng.ColumnHeaders(22).Text = "NILAI SISTEM"
                vBng.ColumnHeaders(23).Text = "KELAS"
                vBng.ColumnHeaders(24).Text = "NJOP_PER_M2"
                vBng.ColumnHeaders(25).Text = "TOTAL NJOP"
                vBng.ColumnHeaders(26).Text = "NJOPTKP"
                vBng.ColumnHeaders(2).Width = 800
                vBng.ColumnHeaders(3).Width = 700
                vBng.ColumnHeaders(4).Width = 700
                vBng.ColumnHeaders(5).Width = 700
                vBng.ColumnHeaders(6).Width = 700
                vBng.ColumnHeaders(7).Width = 700
                vBng.ColumnHeaders(8).Width = 700
                vBng.ColumnHeaders(9).Width = 700
                vBng.ColumnHeaders(10).Width = 700
                vBng.ColumnHeaders(11).Width = 700
                vBng.ColumnHeaders(12).Width = 1200
                vBng.ColumnHeaders(13).Width = 1000
                vBng.ColumnHeaders(14).Width = 700
                vBng.ColumnHeaders(15).Width = 1400
                vBng.ColumnHeaders(16).Width = 700
                vBng.ColumnHeaders(17).Width = 700
                vBng.ColumnHeaders(18).Width = 700
                vBng.ColumnHeaders(19).Width = 700
                vBng.ColumnHeaders(20).Width = 700
                vBng.ColumnHeaders(21).Width = 700
                vBng.ColumnHeaders(22).Width = 1900
                vBng.ColumnHeaders(23).Width = 700
                vBng.ColumnHeaders(24).Width = 1400
                vBng.ColumnHeaders(25).Width = 1900
                vBng.ColumnHeaders(26).Width = 1900
                vBng.ColumnHeaders(3).Alignment = lvwColumnCenter: vBng.ColumnHeaders(4).Alignment = lvwColumnCenter
                vBng.ColumnHeaders(5).Alignment = lvwColumnCenter: vBng.ColumnHeaders(6).Alignment = lvwColumnCenter
                vBng.ColumnHeaders(7).Alignment = lvwColumnCenter: vBng.ColumnHeaders(8).Alignment = lvwColumnCenter
                vBng.ColumnHeaders(9).Alignment = lvwColumnCenter: vBng.ColumnHeaders(10).Alignment = lvwColumnCenter
                vBng.ColumnHeaders(11).Alignment = lvwColumnCenter: vBng.ColumnHeaders(12).Alignment = lvwColumnRight
                vBng.ColumnHeaders(13).Alignment = lvwColumnRight: vBng.ColumnHeaders(14).Alignment = lvwColumnRight
                vBng.ColumnHeaders(15).Alignment = lvwColumnRight: vBng.ColumnHeaders(16).Alignment = lvwColumnCenter
                vBng.ColumnHeaders(17).Alignment = lvwColumnRight
End Sub
Sub judul_nilai_bumi()
On Error Resume Next
vBangunan.ColumnHeaders(3).Text = "PROP"
                vBangunan.ColumnHeaders(4).Text = "KAB"
                vBangunan.ColumnHeaders(5).Text = "KEC"
                vBangunan.ColumnHeaders(6).Text = "KEL"
                vBangunan.ColumnHeaders(7).Text = "BLOK"
                vBangunan.ColumnHeaders(8).Text = "URUT"
                vBangunan.ColumnHeaders(9).Text = "KODE"
                vBangunan.ColumnHeaders(10).Text = "NO BUMI"
                vBangunan.ColumnHeaders(11).Text = "ZNT"
                vBangunan.ColumnHeaders(12).Text = "NIR"
                vBangunan.ColumnHeaders(13).Text = "LUAS"
                vBangunan.ColumnHeaders(14).Text = "JENIS"
                vBangunan.ColumnHeaders(15).Text = "NILAI SISTEM"
                vBangunan.ColumnHeaders(16).Text = "KELAS"
                vBangunan.ColumnHeaders(17).Text = "NJOP BUMI"
                vBangunan.ColumnHeaders(2).Width = 800
                vBangunan.ColumnHeaders(3).Width = 700
                vBangunan.ColumnHeaders(4).Width = 700
                vBangunan.ColumnHeaders(5).Width = 700
                vBangunan.ColumnHeaders(6).Width = 700
                vBangunan.ColumnHeaders(7).Width = 700
                vBangunan.ColumnHeaders(8).Width = 700
                vBangunan.ColumnHeaders(9).Width = 700
                vBangunan.ColumnHeaders(10).Width = 700
                vBangunan.ColumnHeaders(11).Width = 700
                vBangunan.ColumnHeaders(12).Width = 1200
                vBangunan.ColumnHeaders(13).Width = 1000
                vBangunan.ColumnHeaders(14).Width = 700
                vBangunan.ColumnHeaders(15).Width = 1400
                vBangunan.ColumnHeaders(16).Width = 700
                vBangunan.ColumnHeaders(17).Width = 1400
                vBangunan.ColumnHeaders(3).Alignment = lvwColumnCenter: vBangunan.ColumnHeaders(4).Alignment = lvwColumnCenter
                vBangunan.ColumnHeaders(5).Alignment = lvwColumnCenter: vBangunan.ColumnHeaders(6).Alignment = lvwColumnCenter
                vBangunan.ColumnHeaders(7).Alignment = lvwColumnCenter: vBangunan.ColumnHeaders(8).Alignment = lvwColumnCenter
                vBangunan.ColumnHeaders(9).Alignment = lvwColumnCenter: vBangunan.ColumnHeaders(10).Alignment = lvwColumnCenter
                vBangunan.ColumnHeaders(11).Alignment = lvwColumnCenter: vBangunan.ColumnHeaders(12).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(13).Alignment = lvwColumnRight: vBangunan.ColumnHeaders(14).Alignment = lvwColumnRight
                vBangunan.ColumnHeaders(15).Alignment = lvwColumnRight: vBangunan.ColumnHeaders(16).Alignment = lvwColumnCenter
                vBangunan.ColumnHeaders(17).Alignment = lvwColumnRight
End Sub
