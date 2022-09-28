VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmObjek_Pajak_Bg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Objek Pajak Bangunan"
   ClientHeight    =   9045
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   18510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmObjek_Pajak_Bg.frx":0000
   ScaleHeight     =   9045
   ScaleWidth      =   18510
   Begin VB.Frame FF2 
      BorderStyle     =   0  'None
      Caption         =   "[D. DATA TAMBAHAN UNTUK BANGUNAN NON STANDARD]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9360
      Left            =   11115
      TabIndex        =   185
      Top             =   -330
      Width           =   7455
      Begin VB.Frame fDT 
         Caption         =   "Gedung Sekolah (JPB=16)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   60
         TabIndex        =   230
         Top             =   6945
         Width           =   7260
         Begin VB.ComboBox cJPB16 
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
            Left            =   2295
            TabIndex        =   231
            Top             =   225
            Width           =   4890
         End
         Begin VB.Label Label91 
            Caption         =   "Kelas Bangunan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   232
            Top             =   255
            Width           =   1770
         End
      End
      Begin VB.Frame fDT 
         Caption         =   "Olahraga/Rekreasi (JPB=6)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   90
         TabIndex        =   227
         Top             =   2580
         Width           =   7260
         Begin VB.ComboBox cJPB6 
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
            Left            =   2265
            TabIndex        =   228
            Top             =   225
            Width           =   4890
         End
         Begin VB.Label Label90 
            Caption         =   "Kelas Bangunan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   229
            Top             =   255
            Width           =   1770
         End
      End
      Begin VB.Frame fDT 
         Caption         =   "Toko/Apotik/Pasar/Ruko (JPB=4)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   90
         TabIndex        =   224
         Top             =   960
         Width           =   7260
         Begin VB.TextBox JPB4 
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
            Left            =   5550
            TabIndex        =   272
            Top             =   210
            Width           =   1575
         End
         Begin VB.ComboBox cJPB4 
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
            Left            =   1530
            TabIndex        =   225
            Top             =   195
            Width           =   2550
         End
         Begin VB.Label Label104 
            Caption         =   "Luas Bangunan Dengan AC Central"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4140
            TabIndex        =   273
            Top             =   150
            Width           =   1485
         End
         Begin VB.Label Label89 
            Caption         =   "Kelas Bangunan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   226
            Top             =   255
            Width           =   1770
         End
      End
      Begin VB.Frame fDT 
         Caption         =   "Rumah Sakit/Klinik (JPB=5)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   2
         Left            =   90
         TabIndex        =   217
         Top             =   1575
         Width           =   7260
         Begin VB.TextBox JPB5b 
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
            Left            =   5535
            TabIndex        =   220
            Top             =   540
            Width           =   1650
         End
         Begin VB.ComboBox cJPB5 
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
            Left            =   2295
            TabIndex        =   219
            Top             =   225
            Width           =   4890
         End
         Begin VB.TextBox JPB5a 
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
            Left            =   2295
            TabIndex        =   218
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label Label88 
            Caption         =   "Luas Ruangan Lain Dgn AC Central (M2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3990
            TabIndex        =   223
            Top             =   525
            Width           =   1485
         End
         Begin VB.Label Label87 
            Caption         =   "Luas Kamar dgn AC Central"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   135
            TabIndex        =   222
            Top             =   525
            Width           =   1500
         End
         Begin VB.Label Label86 
            Caption         =   "Kelas Bangunan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   221
            Top             =   255
            Width           =   1770
         End
      End
      Begin VB.Frame fDT 
         Caption         =   "Tangki Minyak (JPB=15)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   7
         Left            =   60
         TabIndex        =   212
         Top             =   6345
         Width           =   7260
         Begin VB.TextBox JPB15 
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
            Left            =   1830
            TabIndex        =   214
            Top             =   195
            Width           =   1845
         End
         Begin VB.ComboBox cJPB15 
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
            Left            =   4680
            TabIndex        =   213
            Top             =   195
            Width           =   2520
         End
         Begin VB.Label Label85 
            Caption         =   "Kapasitas Tangki (M3)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   216
            Top             =   255
            Width           =   1770
         End
         Begin VB.Label Label84 
            Caption         =   "Letak Tangki"
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
            Left            =   3735
            TabIndex        =   215
            Top             =   255
            Width           =   960
         End
      End
      Begin VB.Frame fDT 
         Caption         =   "Apartemen (JBP=13)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   6
         Left            =   90
         TabIndex        =   203
         Top             =   5175
         Width           =   7260
         Begin VB.TextBox JPB13d 
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
            Left            =   5325
            TabIndex        =   259
            Top             =   825
            Width           =   1860
         End
         Begin VB.TextBox JPB13b 
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
            Left            =   5325
            TabIndex        =   207
            Top             =   150
            Width           =   1830
         End
         Begin VB.ComboBox cJPB13 
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
            Left            =   1680
            TabIndex        =   206
            Top             =   240
            Width           =   2040
         End
         Begin VB.TextBox JPB13c 
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
            Left            =   5325
            TabIndex        =   205
            Top             =   480
            Width           =   1845
         End
         Begin VB.TextBox JPB13a 
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
            Left            =   1680
            TabIndex        =   204
            Top             =   615
            Width           =   2025
         End
         Begin VB.Label Label102 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Kamar Menggunakan Boiler"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3810
            TabIndex        =   260
            Top             =   780
            Width           =   2010
         End
         Begin VB.Label Label55 
            Caption         =   "Luas Ruangan Lain Dgn AC Central (M2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   180
            TabIndex        =   211
            Top             =   555
            Width           =   1485
         End
         Begin VB.Label Label51 
            Caption         =   "Luas Kamar dgn AC Central"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3840
            TabIndex        =   210
            Top             =   420
            Width           =   1500
         End
         Begin VB.Label Label50 
            Caption         =   "Kelas Bangunan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   180
            TabIndex        =   209
            Top             =   270
            Width           =   1770
         End
         Begin VB.Label Label42 
            Caption         =   "Jumlah Kamar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3840
            TabIndex        =   208
            Top             =   135
            Width           =   1770
         End
      End
      Begin VB.Frame fDT 
         Caption         =   "Bangunan Parkir (JPB=12)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   5
         Left            =   75
         TabIndex        =   200
         Top             =   4560
         Width           =   7260
         Begin VB.ComboBox cJPB12 
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
            Left            =   1635
            TabIndex        =   201
            Top             =   225
            Width           =   5550
         End
         Begin VB.Label Label41 
            Caption         =   "Tipe Bangunan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   202
            Top             =   255
            Width           =   1770
         End
      End
      Begin VB.Frame fDT 
         Caption         =   "Hotel/Wisma (JPB=7)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Index           =   4
         Left            =   90
         TabIndex        =   189
         Top             =   3195
         Width           =   7260
         Begin VB.TextBox JPB7d 
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
            Left            =   5475
            TabIndex        =   257
            Top             =   930
            Width           =   1545
         End
         Begin VB.TextBox JPB7a 
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
            Left            =   1635
            TabIndex        =   194
            Top             =   885
            Width           =   2025
         End
         Begin VB.ComboBox cJPB7b 
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
            Left            =   1635
            TabIndex        =   193
            Top             =   555
            Width           =   2025
         End
         Begin VB.TextBox JPB7c 
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
            Left            =   5475
            TabIndex        =   192
            Top             =   615
            Width           =   1545
         End
         Begin VB.ComboBox cJPB7a 
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
            Left            =   1635
            TabIndex        =   191
            Top             =   225
            Width           =   2040
         End
         Begin VB.TextBox JPB7b 
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
            Left            =   5460
            TabIndex        =   190
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label Label101 
            Caption         =   "Jumlah Kamar Menggunakan Boiler"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3855
            TabIndex        =   258
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label Label18 
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
            Height          =   300
            Left            =   150
            TabIndex        =   199
            Top             =   555
            Width           =   1770
         End
         Begin VB.Label Label58 
            Caption         =   "Jenis Hotel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   198
            Top             =   255
            Width           =   1770
         End
         Begin VB.Label Label57 
            Caption         =   "Luas Kamar dgn AC Central"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3870
            TabIndex        =   197
            Top             =   585
            Width           =   1500
         End
         Begin VB.Label Label56 
            Caption         =   "Luas Ruangan Lain Dgn AC Central (M2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3855
            TabIndex        =   196
            Top             =   165
            Width           =   1485
         End
         Begin VB.Label Label17 
            Caption         =   "Jumlah Kamar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   150
            TabIndex        =   195
            Top             =   900
            Width           =   1500
         End
      End
      Begin VB.Frame fDT 
         Caption         =   "Perkantoran Swasta/Gedung Pemerintah (JPB=2/9)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   90
         TabIndex        =   186
         Top             =   375
         Width           =   7260
         Begin VB.TextBox JPB29 
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
            Left            =   5565
            TabIndex        =   274
            Top             =   225
            Width           =   1575
         End
         Begin VB.ComboBox cJPB29 
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
            TabIndex        =   187
            Top             =   210
            Width           =   2535
         End
         Begin VB.Label Label105 
            Caption         =   "Luas Bangunan Dengan AC Central"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4155
            TabIndex        =   275
            Top             =   165
            Width           =   1485
         End
         Begin VB.Label Label16 
            Caption         =   "Kelas Bangunan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   188
            Top             =   255
            Width           =   1770
         End
      End
   End
   Begin VB.TextBox txtPajak 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Index           =   2
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   265
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame fDT 
      Caption         =   "Tower/Menara Telekomunikasi (JPB=17)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   11100
      TabIndex        =   261
      Top             =   1200
      Width           =   7470
      Begin VB.TextBox cJPB17b 
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
         Left            =   5325
         TabIndex        =   270
         Top             =   195
         Width           =   2100
      End
      Begin VB.TextBox cJPB17 
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
         Left            =   1635
         TabIndex        =   262
         Top             =   195
         Width           =   2745
      End
      Begin VB.Label Label103 
         Caption         =   "Luas Pagar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4440
         TabIndex        =   271
         Top             =   225
         Width           =   1770
      End
      Begin VB.Label Label100 
         Caption         =   "Tinggi Menara (M)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   263
         Top             =   255
         Width           =   1770
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   45
      ScaleHeight     =   300
      ScaleWidth      =   10335
      TabIndex        =   240
      Top             =   2910
      Width           =   10335
      Begin VB.CommandButton cmdProses 
         Caption         =   "&Proses"
         Height          =   330
         Left            =   1605
         TabIndex        =   250
         Top             =   -45
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label LPersen 
         Caption         =   "%Susut"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   255
         Top             =   60
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LEff 
         AutoSize        =   -1  'True
         Caption         =   "Umur Efektif"
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
         Left            =   4785
         TabIndex        =   254
         Top             =   45
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label LFas 
         Caption         =   "Nilai Fasilitas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6915
         TabIndex        =   253
         Top             =   60
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LSusut 
         AutoSize        =   -1  'True
         Caption         =   "Susut"
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
         Left            =   8340
         TabIndex        =   252
         Top             =   45
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label94 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATA FASILITAS"
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
         Left            =   4530
         TabIndex        =   241
         Top             =   60
         Width           =   1260
      End
   End
   Begin VB.CommandButton Command3 
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
      Height          =   435
      Left            =   15510
      TabIndex        =   235
      Top             =   9555
      Width           =   990
   End
   Begin VB.CommandButton Command2 
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
      Height          =   435
      Left            =   16485
      TabIndex        =   234
      Top             =   9555
      Width           =   990
   End
   Begin VB.CommandButton Command1 
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
      Height          =   435
      Left            =   17460
      TabIndex        =   233
      Top             =   9555
      Width           =   990
   End
   Begin VB.Frame FF1 
      Caption         =   "JPB=3/8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   11100
      TabIndex        =   174
      Top             =   -30
      Width           =   7470
      Begin VB.TextBox JPB38 
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
         Left            =   5535
         TabIndex        =   179
         Top             =   180
         Width           =   1650
      End
      Begin VB.TextBox JPB38 
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
         Left            =   2310
         TabIndex        =   178
         Top             =   180
         Width           =   1650
      End
      Begin VB.TextBox JPB38 
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
         Left            =   2310
         TabIndex        =   177
         Top             =   795
         Width           =   1650
      End
      Begin VB.TextBox JPB38 
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
         Left            =   5535
         TabIndex        =   176
         Top             =   495
         Width           =   1650
      End
      Begin VB.TextBox JPB38 
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
         Left            =   2310
         TabIndex        =   175
         Top             =   495
         Width           =   1650
      End
      Begin VB.Label Label15 
         Caption         =   "Lembar Bentang (M)"
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
         TabIndex        =   184
         Top             =   525
         Width           =   1470
      End
      Begin VB.Label Label13 
         Caption         =   "Tinggi Kolom (M)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   183
         Top             =   225
         Width           =   1770
      End
      Begin VB.Label Label10 
         Caption         =   "Luas Mezzanie (M2)"
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
         Left            =   4065
         TabIndex        =   182
         Top             =   525
         Width           =   1500
      End
      Begin VB.Label Label9 
         Caption         =   "Keliling Dinding (M)"
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
         Left            =   4080
         TabIndex        =   181
         Top             =   195
         Width           =   1470
      End
      Begin VB.Label Label2 
         Caption         =   "Daya Dukung Lantai (Kg/M2)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   180
         Top             =   840
         Width           =   2205
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5250
      TabIndex        =   59
      Top             =   8520
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4275
      TabIndex        =   58
      Top             =   8520
      Width           =   990
   End
   Begin VB.CommandButton cmdSave 
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
      Height          =   435
      Left            =   3300
      TabIndex        =   57
      Top             =   8520
      Width           =   990
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   90
      TabIndex        =   84
      Top             =   3105
      Width           =   10290
      Begin VB.Frame Frame18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jumlah Lapangan Tennis"
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
         Height          =   1575
         Left            =   6525
         TabIndex        =   152
         Top             =   135
         Width           =   3705
         Begin VB.TextBox tLap_Tanah2 
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
            Left            =   2445
            TabIndex        =   32
            Top             =   1140
            Width           =   1200
         End
         Begin VB.TextBox tLap_Beton2 
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
            Left            =   2445
            TabIndex        =   28
            Top             =   480
            Width           =   1200
         End
         Begin VB.TextBox tLap_Beton1 
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
            Left            =   1245
            TabIndex        =   27
            Top             =   480
            Width           =   1185
         End
         Begin VB.TextBox tLap_Aspal2 
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
            Left            =   2445
            TabIndex        =   30
            Top             =   810
            Width           =   1200
         End
         Begin VB.TextBox tLap_Tanah1 
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
            Left            =   1245
            TabIndex        =   31
            Top             =   1140
            Width           =   1185
         End
         Begin VB.TextBox tLap_Aspal1 
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
            Left            =   1245
            TabIndex        =   29
            Top             =   810
            Width           =   1185
         End
         Begin VB.Label Label63 
            BackStyle       =   0  'Transparent
            Caption         =   "Beton"
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
            TabIndex        =   157
            Top             =   510
            Width           =   465
         End
         Begin VB.Label Label62 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanah/Rumput"
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
            Left            =   135
            TabIndex        =   156
            Top             =   1230
            Width           =   1500
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Aspal"
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
            Left            =   135
            TabIndex        =   155
            Top             =   855
            Width           =   1470
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dengan Lampu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1275
            TabIndex        =   154
            Top             =   225
            Width           =   1065
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanpa   Lampu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2490
            TabIndex        =   153
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pemadam Kebakaran"
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
         Height          =   1230
         Left            =   6495
         TabIndex        =   168
         Top             =   1710
         Width           =   3735
         Begin VB.ComboBox cboPajak 
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
            Index           =   15
            Left            =   1635
            TabIndex        =   39
            Text            =   "Tidak Ada"
            Top             =   855
            Width           =   1980
         End
         Begin VB.ComboBox cboPajak 
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
            Index           =   14
            Left            =   1635
            TabIndex        =   38
            Text            =   "Tidak Ada"
            Top             =   540
            Width           =   1980
         End
         Begin VB.ComboBox cboPajak 
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
            Index           =   13
            Left            =   1635
            TabIndex        =   37
            Text            =   "Tidak Ada"
            Top             =   225
            Width           =   1980
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "1. Hydran"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   105
            TabIndex        =   171
            Top             =   255
            Width           =   885
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "2. Sprinkler"
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
            Left            =   120
            TabIndex        =   170
            Top             =   585
            Width           =   1500
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "3. Fire Alarm"
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
            Left            =   120
            TabIndex        =   169
            Top             =   885
            Width           =   1500
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lebar Tangga Berjalan"
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
         Height          =   930
         Left            =   3285
         TabIndex        =   165
         Top             =   1710
         Width           =   3195
         Begin VB.TextBox tTangga2 
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
            Left            =   1605
            TabIndex        =   36
            Top             =   555
            Width           =   1500
         End
         Begin VB.TextBox tTangga1 
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
            Left            =   1605
            TabIndex        =   35
            Top             =   225
            Width           =   1500
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   ">0.80 M"
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
            TabIndex        =   167
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "<=0.80 M"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   166
            Top             =   255
            Width           =   1215
         End
      End
      Begin VB.Frame Frame16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pagar"
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
         Height          =   930
         Left            =   75
         TabIndex        =   162
         Top             =   1710
         Width           =   3195
         Begin VB.ComboBox cboPajak 
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
            Index           =   16
            Left            =   1620
            TabIndex        =   33
            Top             =   180
            Width           =   1500
         End
         Begin VB.TextBox tPagar 
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
            Left            =   1620
            TabIndex        =   34
            Top             =   540
            Width           =   1500
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Bahan Pagar"
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
            Left            =   180
            TabIndex        =   164
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "Panjang Pagar (M)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   180
            TabIndex        =   163
            Top             =   570
            Width           =   1350
         End
      End
      Begin VB.Frame Frame24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jumlah Lift"
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
         Height          =   1245
         Left            =   75
         TabIndex        =   158
         Top             =   2640
         Width           =   3210
         Begin VB.TextBox tLift1 
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
            Left            =   1605
            TabIndex        =   40
            Top             =   225
            Width           =   1500
         End
         Begin VB.TextBox tLift2 
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
            Left            =   1605
            TabIndex        =   41
            Top             =   555
            Width           =   1500
         End
         Begin VB.TextBox tLift3 
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
            Left            =   1605
            TabIndex        =   42
            Top             =   885
            Width           =   1500
         End
         Begin VB.Label Label81 
            BackStyle       =   0  'Transparent
            Caption         =   "Penumpang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   225
            TabIndex        =   161
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label80 
            BackStyle       =   0  'Transparent
            Caption         =   "Kapsul"
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
            Left            =   240
            TabIndex        =   160
            Top             =   570
            Width           =   825
         End
         Begin VB.Label Label79 
            BackStyle       =   0  'Transparent
            Caption         =   "Barang"
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
            Left            =   240
            TabIndex        =   159
            Top             =   900
            Width           =   1485
         End
      End
      Begin VB.Frame Frame25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kolam Renang"
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
         Height          =   930
         Left            =   6495
         TabIndex        =   149
         Top             =   2940
         Width           =   3735
         Begin VB.TextBox tKolam 
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
            Left            =   1635
            TabIndex        =   47
            Top             =   555
            Width           =   1995
         End
         Begin VB.ComboBox cboPajak 
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
            Index           =   19
            Left            =   1635
            TabIndex        =   46
            Text            =   "Diplester"
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label83 
            BackStyle       =   0  'Transparent
            Caption         =   "Luas (M2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   195
            TabIndex        =   151
            Top             =   600
            Width           =   1560
         End
         Begin VB.Label Label82 
            BackStyle       =   0  'Transparent
            Caption         =   "Finishing Kolam"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   195
            TabIndex        =   150
            Top             =   255
            Width           =   1230
         End
      End
      Begin VB.Frame Frame19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Listrik dan AC"
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
         Height          =   1575
         Left            =   90
         TabIndex        =   144
         Top             =   135
         Width           =   3195
         Begin VB.TextBox tListrik 
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
            Left            =   1635
            TabIndex        =   19
            Top             =   210
            Width           =   1500
         End
         Begin VB.ComboBox cboPajak 
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
            Index           =   17
            Left            =   1635
            TabIndex        =   22
            Text            =   "Tidak Ada"
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox tAC2 
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
            Left            =   1635
            TabIndex        =   21
            Top             =   870
            Width           =   1500
         End
         Begin VB.TextBox tAC1 
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
            Left            =   1635
            TabIndex        =   20
            Top             =   540
            Width           =   1500
         End
         Begin VB.Label Label67 
            BackStyle       =   0  'Transparent
            Caption         =   "Daya Listrik (Watt)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   135
            TabIndex        =   148
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label66 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah AC Window"
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
            Left            =   120
            TabIndex        =   147
            Top             =   915
            Width           =   1620
         End
         Begin VB.Label Label65 
            BackStyle       =   0  'Transparent
            Caption         =   "AC Central"
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
            Left            =   135
            TabIndex        =   146
            Top             =   1215
            Width           =   1170
         End
         Begin VB.Label Label64 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Ac Split"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   145
            Top             =   570
            Width           =   1215
         End
      End
      Begin VB.Frame Frame23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Luas Perkerasan Halaman (M2)"
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
         Height          =   1575
         Left            =   3300
         TabIndex        =   139
         Top             =   135
         Width           =   3210
         Begin VB.TextBox tHal4 
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
            Left            =   1605
            TabIndex        =   26
            Top             =   1185
            Width           =   1500
         End
         Begin VB.TextBox tHal3 
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
            Left            =   1605
            TabIndex        =   25
            Top             =   855
            Width           =   1500
         End
         Begin VB.TextBox tHal2 
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
            Left            =   1605
            TabIndex        =   24
            Top             =   525
            Width           =   1500
         End
         Begin VB.TextBox tHal1 
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
            Left            =   1605
            TabIndex        =   23
            Top             =   195
            Width           =   1500
         End
         Begin VB.Label Label78 
            BackStyle       =   0  'Transparent
            Caption         =   "Penutup Lantai"
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
            Left            =   210
            TabIndex        =   143
            Top             =   1230
            Width           =   1500
         End
         Begin VB.Label Label77 
            BackStyle       =   0  'Transparent
            Caption         =   "Berat"
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
            Left            =   210
            TabIndex        =   142
            Top             =   915
            Width           =   1470
         End
         Begin VB.Label Label76 
            BackStyle       =   0  'Transparent
            Caption         =   "Sedang"
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
            Left            =   195
            TabIndex        =   141
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label75 
            BackStyle       =   0  'Transparent
            Caption         =   "Ringan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   140
            Top             =   255
            Width           =   1215
         End
      End
      Begin VB.Frame Frame22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1245
         Left            =   3300
         TabIndex        =   136
         Top             =   2625
         Width           =   3180
         Begin VB.TextBox tGenset 
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
            Left            =   1605
            TabIndex        =   45
            Top             =   855
            Width           =   1500
         End
         Begin VB.TextBox tPABX 
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
            Left            =   1605
            TabIndex        =   43
            Top             =   225
            Width           =   1500
         End
         Begin VB.TextBox tSumur 
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
            Left            =   1605
            TabIndex        =   44
            Top             =   540
            Width           =   1500
         End
         Begin VB.Label Label99 
            BackStyle       =   0  'Transparent
            Caption         =   "Kapasitas Genset"
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
            Left            =   90
            TabIndex        =   256
            Top             =   900
            Width           =   1860
         End
         Begin VB.Label Label73 
            BackStyle       =   0  'Transparent
            Caption         =   "Dalam Sumur Artesis"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   90
            TabIndex        =   138
            Top             =   585
            Width           =   1500
         End
         Begin VB.Label Label74 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Saluran PABX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   75
            TabIndex        =   137
            Top             =   285
            Width           =   1515
         End
      End
   End
   Begin VB.Frame Frame12 
      Height          =   450
      Left            =   75
      TabIndex        =   266
      Top             =   6945
      Width           =   10275
      Begin VB.TextBox tKet 
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
         Left            =   960
         TabIndex        =   49
         Text            =   "-"
         Top             =   120
         Width           =   4455
      End
      Begin VB.TextBox txtPajak 
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
         Left            =   3525
         TabIndex        =   48
         Top             =   90
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.TextBox txtPajak 
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
         Index           =   6
         Left            =   6855
         TabIndex        =   50
         Top             =   105
         Width           =   3375
      End
      Begin VB.CommandButton CMDHITUNG 
         BackColor       =   &H00FFFF00&
         Caption         =   "Proses Nilai Sistem>>"
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   267
         Top             =   0
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label106 
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
         Height          =   285
         Left            =   120
         TabIndex        =   280
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label97 
         BackStyle       =   0  'Transparent
         Caption         =   "Nilai Sistem"
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
         Left            =   2400
         TabIndex        =   269
         Top             =   165
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label98 
         BackStyle       =   0  'Transparent
         Caption         =   "Nilai Individual"
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
         Left            =   5520
         TabIndex        =   268
         Top             =   165
         Width           =   1755
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9435
      Left            =   10320
      ScaleHeight     =   9435
      ScaleWidth      =   8535
      TabIndex        =   278
      Top             =   30
      Width           =   8535
      Begin VB.Image Image2 
         Height          =   9015
         Left            =   180
         Picture         =   "frmObjek_Pajak_Bg.frx":0BD4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   18765
      TabIndex        =   242
      Top             =   7410
      Width           =   18765
      Begin VB.Label Label95 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATA PETUGAS"
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
         Left            =   4560
         TabIndex        =   243
         Top             =   45
         Width           =   1215
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   90
      TabIndex        =   77
      Top             =   7575
      Width           =   10275
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
         Index           =   7
         Left            =   1335
         MaxLength       =   30
         TabIndex        =   52
         Top             =   420
         Width           =   1995
      End
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
         Index           =   8
         Left            =   4695
         MaxLength       =   30
         TabIndex        =   54
         Top             =   420
         Width           =   2085
      End
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
         Index           =   9
         Left            =   8040
         MaxLength       =   30
         TabIndex        =   56
         Top             =   420
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker dtPajak 
         Height          =   300
         Index           =   1
         Left            =   3330
         TabIndex        =   53
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   152174593
         CurrentDate     =   41486
      End
      Begin MSComCtl2.DTPicker dtPajak 
         Height          =   300
         Index           =   2
         Left            =   6720
         TabIndex        =   55
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   152174593
         CurrentDate     =   41486
      End
      Begin MSComCtl2.DTPicker dtPajak 
         Height          =   300
         Index           =   0
         Left            =   135
         TabIndex        =   51
         Top             =   420
         Width           =   1245
         _ExtentX        =   2196
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
         Format          =   152174593
         CurrentDate     =   41486
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1770
         TabIndex        =   83
         Top             =   165
         Width           =   1260
      End
      Begin VB.Label Label37 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3450
         TabIndex        =   82
         Top             =   165
         Width           =   1215
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5190
         TabIndex        =   81
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label35 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6795
         TabIndex        =   80
         Top             =   165
         Width           =   1215
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
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   8490
         TabIndex        =   79
         Top             =   180
         Width           =   1215
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   78
         Top             =   165
         Width           =   1260
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   15
      ScaleHeight     =   300
      ScaleWidth      =   18825
      TabIndex        =   236
      Top             =   -15
      Width           =   18825
      Begin VB.TextBox txtPajak 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   15
         TabIndex        =   276
         Top             =   30
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.Label Label92 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JENIS TRANSAKSI DAN NOMOR FORMULIR"
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
         Left            =   3615
         TabIndex        =   237
         Top             =   60
         Width           =   3165
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   495
      Left            =   90
      TabIndex        =   61
      Top             =   195
      Width           =   10260
      Begin VB.CheckBox chPajak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Penilaian Individual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   7635
         TabIndex        =   3
         Top             =   210
         Width           =   1815
      End
      Begin VB.Frame Frame21 
         Caption         =   "Kolam Renang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   131
         Top             =   5310
         Width           =   8265
         Begin VB.TextBox Text44 
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
            Left            =   1620
            TabIndex        =   133
            Top             =   210
            Width           =   2430
         End
         Begin VB.ComboBox cboPajak 
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
            Index           =   18
            Left            =   5730
            TabIndex        =   132
            Top             =   195
            Width           =   2430
         End
         Begin VB.Label Label71 
            Caption         =   "Luas (M2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   195
            TabIndex        =   135
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label Label72 
            Caption         =   "Finishing Kolam"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4290
            TabIndex        =   134
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Jumlah Lift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   4140
         TabIndex        =   124
         Top             =   1590
         Width           =   4125
         Begin VB.TextBox Text41 
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
            Left            =   1635
            TabIndex        =   127
            Top             =   270
            Width           =   2430
         End
         Begin VB.TextBox Text42 
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
            Left            =   1635
            TabIndex        =   126
            Top             =   645
            Width           =   2430
         End
         Begin VB.TextBox Text43 
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
            Left            =   1620
            TabIndex        =   125
            Top             =   1035
            Width           =   2430
         End
         Begin VB.Label Label68 
            Caption         =   "Penumpang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   225
            TabIndex        =   130
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label69 
            Caption         =   "Kapsul"
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
            Left            =   240
            TabIndex        =   129
            Top             =   660
            Width           =   825
         End
         Begin VB.Label Label70 
            Caption         =   "Barang"
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
            Left            =   255
            TabIndex        =   128
            Top             =   1065
            Width           =   1485
         End
      End
      Begin VB.Frame Frame14 
         Height          =   1260
         Left            =   4140
         TabIndex        =   119
         Top             =   4080
         Width           =   4125
         Begin VB.TextBox Text6 
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
            Left            =   1575
            TabIndex        =   121
            Top             =   720
            Width           =   2430
         End
         Begin VB.TextBox Text14 
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
            Left            =   1590
            TabIndex        =   120
            Top             =   285
            Width           =   2430
         End
         Begin VB.Label Label38 
            Caption         =   "Jumlah Saluran PES. PABX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   210
            TabIndex        =   123
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label Label28 
            Caption         =   "Kedalaman Sumur Artesis (M)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   180
            TabIndex        =   122
            Top             =   660
            Width           =   1275
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Kolam Renang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         TabIndex        =   114
         Top             =   5220
         Width           =   8265
         Begin VB.ComboBox cboPajak 
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
            Index           =   11
            Left            =   5730
            TabIndex        =   116
            Top             =   195
            Width           =   2430
         End
         Begin VB.TextBox Text26 
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
            Left            =   1620
            TabIndex        =   115
            Top             =   210
            Width           =   2430
         End
         Begin VB.Label Label52 
            Caption         =   "Finishing Kolam"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4290
            TabIndex        =   118
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label Label14 
            Caption         =   "Luas (M2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   195
            TabIndex        =   117
            Top             =   240
            Width           =   1560
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Jumlah Lift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   4140
         TabIndex        =   107
         Top             =   1590
         Width           =   4125
         Begin VB.TextBox Text20 
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
            Left            =   1620
            TabIndex        =   110
            Top             =   1035
            Width           =   2430
         End
         Begin VB.TextBox Text21 
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
            Left            =   1635
            TabIndex        =   109
            Top             =   645
            Width           =   2430
         End
         Begin VB.TextBox Text22 
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
            Left            =   1635
            TabIndex        =   108
            Top             =   270
            Width           =   2430
         End
         Begin VB.Label Label24 
            Caption         =   "Barang"
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
            Left            =   255
            TabIndex        =   113
            Top             =   1065
            Width           =   1485
         End
         Begin VB.Label Label29 
            Caption         =   "Kapsul"
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
            Left            =   240
            TabIndex        =   112
            Top             =   660
            Width           =   825
         End
         Begin VB.Label Label23 
            Caption         =   "Penumpang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   225
            TabIndex        =   111
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Jumlah Lapangan Tennis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   0
         TabIndex        =   95
         Top             =   1590
         Width           =   4125
         Begin VB.TextBox Text7 
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
            Left            =   2850
            TabIndex        =   101
            Top             =   1065
            Width           =   1200
         End
         Begin VB.TextBox Text9 
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
            Left            =   2850
            TabIndex        =   100
            Top             =   420
            Width           =   1200
         End
         Begin VB.TextBox Text10 
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
            Left            =   1650
            TabIndex        =   99
            Top             =   420
            Width           =   1185
         End
         Begin VB.TextBox Text11 
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
            Left            =   2850
            TabIndex        =   98
            Top             =   750
            Width           =   1200
         End
         Begin VB.TextBox Text12 
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
            Left            =   1650
            TabIndex        =   97
            Top             =   1065
            Width           =   1185
         End
         Begin VB.TextBox Text28 
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
            Left            =   1650
            TabIndex        =   96
            Top             =   750
            Width           =   1185
         End
         Begin VB.Label Label30 
            Caption         =   "Beton"
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
            TabIndex        =   106
            Top             =   450
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Tanah Liat/Rumput"
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
            Left            =   135
            TabIndex        =   105
            Top             =   1110
            Width           =   1500
         End
         Begin VB.Label Label3 
            Caption         =   "Aspal"
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
            Left            =   150
            TabIndex        =   104
            Top             =   780
            Width           =   1470
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Dengan Lampu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   1725
            TabIndex        =   103
            Top             =   195
            Width           =   1065
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Tanpa   Lampu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   2940
            TabIndex        =   102
            Top             =   195
            Width           =   1050
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Pagar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   4155
         TabIndex        =   90
         Top             =   3165
         Width           =   4125
         Begin VB.TextBox Text32 
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
            Left            =   1605
            TabIndex        =   92
            Top             =   210
            Width           =   2430
         End
         Begin VB.ComboBox cboPajak 
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
            Index           =   12
            Left            =   1605
            TabIndex        =   91
            Top             =   555
            Width           =   2430
         End
         Begin VB.Label Label20 
            Caption         =   "Panjang Pagar (M)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   180
            TabIndex        =   94
            Top             =   255
            Width           =   1350
         End
         Begin VB.Label Label26 
            Caption         =   "Bahan Pagar"
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
            Left            =   195
            TabIndex        =   93
            Top             =   585
            Width           =   1170
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Lebar Tangga Berjalan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   0
         TabIndex        =   85
         Top             =   3165
         Width           =   4125
         Begin VB.TextBox Text31 
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
            Left            =   1635
            TabIndex        =   87
            Top             =   555
            Width           =   2430
         End
         Begin VB.TextBox Text30 
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
            Left            =   1635
            TabIndex        =   86
            Top             =   240
            Width           =   2430
         End
         Begin VB.Label Label49 
            Caption         =   ">0.80 M"
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
            Left            =   210
            TabIndex        =   89
            Top             =   585
            Width           =   1485
         End
         Begin VB.Label Label25 
            Caption         =   "<=0.80 M"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   165
            TabIndex        =   88
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.CheckBox chPajak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   0
         Top             =   195
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox chPajak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2610
         TabIndex        =   1
         Top             =   210
         Width           =   1815
      End
      Begin VB.CheckBox chPajak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   5280
         TabIndex        =   2
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   4365
      TabIndex        =   172
      Top             =   600
      Width           =   5985
      Begin VB.ComboBox cTPajak 
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
         Left            =   4665
         TabIndex        =   6
         Top             =   165
         Width           =   1155
      End
      Begin VB.TextBox txtPajak 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   0
         Left            =   1155
         TabIndex        =   5
         Top             =   165
         Width           =   2385
      End
      Begin VB.Label Label96 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO. LSPOP"
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
         TabIndex        =   251
         Top             =   225
         Width           =   810
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
         Height          =   195
         Left            =   3585
         TabIndex        =   173
         Top             =   225
         Width           =   1755
      End
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   570
      Left            =   90
      TabIndex        =   62
      Top             =   585
      Width           =   4335
      Begin VB.CommandButton cmdNOP 
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
         Left            =   3885
         Picture         =   "frmObjek_Pajak_Bg.frx":403D
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   165
         Width           =   330
      End
      Begin MSMask.MaskEdBox aNOP 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   165
         Width           =   3150
         _ExtentX        =   5556
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
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "N.O.P"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   150
         TabIndex        =   63
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   30
      ScaleHeight     =   300
      ScaleWidth      =   10365
      TabIndex        =   238
      Top             =   1170
      Width           =   10365
      Begin VB.CheckBox cStandard 
         BackColor       =   &H00404080&
         Caption         =   "Non Standard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   8985
         TabIndex        =   264
         Top             =   30
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label LLangit2 
         Caption         =   "Langit2"
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
         Left            =   8100
         TabIndex        =   249
         Top             =   15
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label LLantai 
         Caption         =   "Lantai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   248
         Top             =   15
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label LDinding 
         Caption         =   "Dinding"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6720
         TabIndex        =   247
         Top             =   15
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label LAtap 
         Caption         =   "Atap"
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
         Left            =   6195
         TabIndex        =   246
         Top             =   0
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label LDBKB 
         Caption         =   "DBKB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4575
         TabIndex        =   245
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label LTipe 
         Caption         =   "Tipe Bangunan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3150
         TabIndex        =   244
         Top             =   -15
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label93 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RINCIAN DATA BANGUNAN"
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
         Left            =   4125
         TabIndex        =   239
         Top             =   60
         Width           =   2010
      End
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1560
      Left            =   105
      TabIndex        =   64
      Top             =   1365
      Width           =   10245
      Begin VB.ComboBox cboPajak 
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
         Index           =   0
         Left            =   1395
         TabIndex        =   10
         Top             =   510
         Width           =   4245
      End
      Begin VB.CheckBox chEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5520
         TabIndex        =   279
         Top             =   525
         Width           =   300
      End
      Begin MSComCtl2.UpDown xUP 
         Height          =   315
         Left            =   1965
         TabIndex        =   277
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboPajak 
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
         Index           =   1
         Left            =   1395
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Top             =   165
         Width           =   585
      End
      Begin VB.ComboBox cboPajak 
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
         Index           =   9
         Left            =   7500
         TabIndex        =   18
         Top             =   1155
         Width           =   2625
      End
      Begin VB.ComboBox cboPajak 
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
         Index           =   8
         Left            =   7500
         TabIndex        =   17
         Top             =   825
         Width           =   2625
      End
      Begin VB.ComboBox cboPajak 
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
         Index           =   7
         Left            =   7500
         TabIndex        =   16
         Top             =   495
         Width           =   2625
      End
      Begin VB.ComboBox cboPajak 
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
         Index           =   6
         Left            =   7500
         TabIndex        =   15
         Top             =   165
         Width           =   2625
      End
      Begin VB.ComboBox cboPajak 
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
         Index           =   5
         Left            =   4335
         TabIndex        =   14
         Top             =   1170
         Width           =   1500
      End
      Begin VB.ComboBox cboPajak 
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
         Index           =   4
         Left            =   4335
         TabIndex        =   13
         Top             =   840
         Width           =   1500
      End
      Begin VB.ComboBox cboPajak 
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
         Index           =   3
         Left            =   1395
         TabIndex        =   12
         Top             =   1170
         Width           =   1500
      End
      Begin VB.ComboBox cboPajak 
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
         Index           =   2
         Left            =   1395
         TabIndex        =   11
         Top             =   840
         Width           =   1500
      End
      Begin VB.TextBox txtPajak 
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
         Index           =   4
         Left            =   2790
         TabIndex        =   8
         Top             =   165
         Width           =   720
      End
      Begin VB.TextBox txtPajak 
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
         Index           =   3
         Left            =   4740
         MaxLength       =   4
         TabIndex        =   9
         Top             =   180
         Width           =   1110
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   5925
         X2              =   5925
         Y1              =   105
         Y2              =   2565
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Luas (M2)"
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
         Left            =   3525
         TabIndex        =   76
         Top             =   225
         Width           =   1155
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jum Lt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   75
         Top             =   225
         Width           =   705
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Lantai"
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
         Left            =   6120
         TabIndex        =   74
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label cbLangit 
         BackStyle       =   0  'Transparent
         Caption         =   "Langit-Langit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   73
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Bangunan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   72
         Top             =   570
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kondisi Bangunan"
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
         Left            =   3015
         TabIndex        =   71
         Top             =   930
         Width           =   1260
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Dinding"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   70
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Bangunan Ke-"
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
         TabIndex        =   69
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Dibangun"
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
         Left            =   135
         TabIndex        =   68
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Renovasi"
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
         TabIndex        =   67
         Top             =   1245
         Width           =   1485
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Atap"
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
         Left            =   6105
         TabIndex        =   66
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Konstruksi"
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
         Left            =   3030
         TabIndex        =   65
         Top             =   1260
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   9195
      Left            =   -30
      Picture         =   "frmObjek_Pajak_Bg.frx":4D07
      Stretch         =   -1  'True
      Top             =   -135
      Width           =   10365
   End
End
Attribute VB_Name = "frmObjek_Pajak_Bg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cTipe, JPB
Dim xKode, xJPB
Dim xDinding, xLantai, xAtap, xLangit2
Dim umur_EFF
Dim FListrik, FAC1, FAC2, FAC3, nDBKB, xSUSUT
'Dim DAYA_LISTRIK, JUM_SPLIT, JUM_WINDOW
'Dim LUAS_HRINGAN, LUAS_HSEDANG, LUAS_HBERAT, LUAS_HPENUTUP
'Dim JUM_LAP_BETON1, JUM_LAP_BETON2, JUM_LAP_ASPAL1, JUM_LAP_ASPAL2, JUM_LAP_RUMPUT1, JUM_LAP_RUMPUT2, JUM_LAP_BETON11, JUM_LAP_BETON21, JUM_LAP_ASPAL11, JUM_LAP_ASPAL21, JUM_LAP_RUMPUT11, JUM_LAP_RUMPUT21
'Dim PANJANG_PAGAR, BAHAN_PAGAR1, BAHAN_PAGAR2, LEBAR_TANGGA1, LEBAR_TANGGA2, JUM_LIFT1, JUM_LIFT2, JUM_LIFT3, JUM_PABX, DALAM_SUMUR
'Dim BAKAR_H, BAKAR_S, BAKAR_F
'Dim JLIFT(100), Nil_AC_Central(10)
'Dim nSistem, Luas_Kolam, JUM_AC_CENTRAL, JUM_GENSET, Nil_Boiler_Ht, Nil_Boiler_Ap, nMezanin, nDUKUNG
Dim totChar, xTB
Dim t_Luas, t_Nilai, t_NJOP

Private Sub cmdBangunan_Click()
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
frmOP_Tanah.Show
End Sub


Private Sub aNOP_Change()
On Error Resume Next
txtPajak(1).Text = aNOP.Text
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
On Error GoTo Salah
cmdNOP_Click
chEdit.Value = 0
If chPajak(1).Value = 1 Or chPajak(2).Value = 1 Then
    'Panggil Data Bangunan
    cboPajak(1).Text = 1
    StrQ1 = "Select * From DAT_OP_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
    openDB (StrQ1)
    call_data
Else
    StrQ1 = "Select * From DAT_OP_BUMI WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_URUT ASC"
    openDB (StrQ1)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then
        MsgBox "Nomor Objerk Pajak tidak terdaftar...", vbCritical, "Tetnong...!"
    End If
End If
Me.Width = 10395
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cboPajak_Click(Index As Integer)
On Error GoTo Salah
Select Case Index
Case 0
    If cboPajak(0).ListIndex = 13 Then
        Label12.Caption = "Luas Kanopi"
    Else
        Label12.Caption = "Luas (M2)"
    End If
    'Case 1, 10, 11, 14
    xJPB = Left(cboPajak(0).Text, 2) * 1
    'If (xJPB = 1 Or xJPB = 3 Or xJPB = 8 Or xJPB = 10 Or xJPB = 11) Or ((xJPB = 2 Or xJPB = 4 Or xJPB = 5 Or xJPB = 7 Or xJPB = 9) And cStandard.Value = 0 And txtPajak(4).Text <= 4) Then
    If (xJPB = 1 Or xJPB = 10 Or xJPB = 11 Or xJPB = 14) Or ((xJPB = 2 Or xJPB = 4 Or xJPB = 5 Or xJPB = 7 Or xJPB = 9) And cStandard.Value = 0 And txtPajak(4).Text <= 4) Then
    'If (cboPajak(0).ListIndex = 0 Or cboPajak(0).ListIndex = 9 Or cboPajak(0).ListIndex = 10 Or cboPajak(0).ListIndex = 13) Or txtPajak(4).Text > 4 Then
        
    Else
        cStandard.Value = 1
        xxNon = cboPajak(0).ListIndex + 1
        cStandard.Value = 0
        frmTambahan.Show vbModal
    End If
 Case 1
    If chPajak(1).Value = 1 Or chPajak(2).Value = 1 Then
        StrQ = "Select * From DAT_OP_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' " 'and (NO_BNG='" & cboPajak(1).Text * 1 & "') ORDER BY NO_BNG ASC "
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            If rPajak!NO_BNG * 1 = cboPajak(1).Text * 1 Then
                txtPajak(0).Text = rPajak!NO_FORMULIR_LSPOP
            End If
            rPajak.MoveNext
        Loop
    End If
'Case 3, 9
'Case 4
'Case 5
'Case 11
'Case 12
Case 13, 14, 15
'    If cboPajak(Index).Text = cboPajak(Index).List(0) Then
'        tBakar1.Locked = True
'        tBakar1.Enabled = False
'    Else
'        tBakar1.Locked = False
'        tBakar1.Enabled = True
'    End If
Case 16
   
Case 19
    
End Select
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cboPajak_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
    Case 0, 4 To 9, 13 To 19
        If InStr("0123456789-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
         KeyAscii = 0
        End If
    Case 1 To 3
        If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
         KeyAscii = 0
        End If
End Select

End Sub

Private Sub cboPajak_LostFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
     
     cboPajak(Index).Text = cboPajak(Index).List(cboPajak(Index).Text - 1)
     If cboPajak(Index).Text = "" Then cboPajak(Index).Text = cboPajak(Index).List(0)
    xJPB = Left(Trim(cboPajak(Index).Text), 2) * 1
    If cboPajak(Index).ListIndex = 13 Then
        Label12.Caption = "Luas Kanopi"
    Else
        Label12.Caption = "Luas (M2)"
    End If
    'Case 1, 10, 11, 14
    
    
    'If (xJPB = 1 Or xJPB = 3 Or xJPB = 8 Or xJPB = 10 Or xJPB = 11) Or ((xJPB = 2 Or xJPB = 4 Or xJPB = 5 Or xJPB = 7 Or xJPB = 9) And cStandard.Value = 0 And txtPajak(4).Text <= 4) Then
    If (xJPB = 1 Or xJPB = 10 Or xJPB = 11 Or xJPB = 14) Or ((xJPB = 2 Or xJPB = 4 Or xJPB = 5 Or xJPB = 7 Or xJPB = 9) And cStandard.Value = 0 And txtPajak(4).Text <= 4) Then
    'If (cboPajak(0).ListIndex = 0 Or cboPajak(0).ListIndex = 9 Or cboPajak(0).ListIndex = 10 Or cboPajak(0).ListIndex = 13) Or txtPajak(4).Text > 4 Then
        
    Else
        cStandard.Value = 1
        xxNon = xJPB 'cboPajak(Index).ListIndex + 1
        cStandard.Value = 0
        frmTambahan.Show vbModal
    End If
Case 1 'To 2
    If cboPajak(Index).Text = "" Or cboPajak(Index).Text = 0 Then cboPajak(Index).Text = 1 'cboPajak(Index).List(0)
Case 2
    For i = 0 To cboPajak(Index).ListCount - 1
        If (UCase(cboPajak(Index).List(i)) Like "*" + UCase(cboPajak(Index).Text) + "*" = True) Then
            cboPajak(Index).Text = cboPajak(Index).List(i)
            Exit Sub
        End If
          If i = cboPajak(Index).ListCount - 1 Then
            If UCase(cboPajak(Index).List(i)) Like "*" + UCase(cboPajak(Index).Text) + "*" = False Then
                cboPajak(Index).Text = cboPajak(Index).List(0)
                Exit Sub
            End If
        End If
    Next
    If cboPajak(2).Text * 1 > cTPajak.Text * 1 Then
        cboPajak(2).Text = cTPajak.Text
    End If
Case 3
     For i = 0 To cboPajak(Index).ListCount - 1
        If (UCase(cboPajak(Index).List(i)) Like "*" + UCase(cboPajak(Index).Text) + "*" = True) Then
            cboPajak(Index).Text = cboPajak(Index).List(i)
            Exit Sub
        End If
          If i = cboPajak(Index).ListCount - 1 Then
            If UCase(cboPajak(Index).List(i)) Like "*" + UCase(cboPajak(Index).Text) + "*" = False Then
                cboPajak(Index).Text = cboPajak(Index).List(0)
                Exit Sub
            End If
        End If
    Next
    If cboPajak(Index).Text = "" Then cboPajak(Index).Text = 0
    If cboPajak(3).Text < cboPajak(2).Text Then
        cboPajak(3).Text = 0
    End If
Case 4, 5
    cboPajak(Index).Text = cboPajak(Index).List(cboPajak(Index).Text - 1)
    If cboPajak(Index).Text <> cboPajak(Index).List(0) And cboPajak(Index).Text <> cboPajak(Index).List(1) And cboPajak(Index).Text <> cboPajak(Index).List(2) And cboPajak(Index).Text <> cboPajak(Index).List(3) Then
        cboPajak(Index).Text = cboPajak(Index).List(1)
    End If
Case 6, 8
    cboPajak(Index).Text = cboPajak(Index).List(cboPajak(Index).Text - 1)
        If cboPajak(Index).Text <> cboPajak(Index).List(0) And cboPajak(Index).Text <> cboPajak(Index).List(1) And cboPajak(Index).Text <> cboPajak(Index).List(2) And cboPajak(Index).Text <> cboPajak(Index).List(3) Then 'And cboPajak(Index).Text <> cboPajak(Index).List(4) And cboPajak(Index).Text <> cboPajak(Index).List(5) Then
            cboPajak(Index).Text = cboPajak(Index).List(4)
        End If
Case 7
        cboPajak(Index).Text = cboPajak(Index).List(cboPajak(Index).Text - 1)
        If cboPajak(Index).Text <> cboPajak(Index).List(0) And cboPajak(Index).Text <> cboPajak(Index).List(1) And cboPajak(Index).Text <> cboPajak(Index).List(2) And cboPajak(Index).Text <> cboPajak(Index).List(3) And cboPajak(Index).Text <> cboPajak(Index).List(4) And cboPajak(Index).Text <> cboPajak(Index).List(5) And cboPajak(Index).Text <> cboPajak(Index).List(6) Then
            cboPajak(Index).Text = cboPajak(Index).List(1)
        End If
Case 9
    cboPajak(Index).Text = cboPajak(Index).List(cboPajak(Index).Text - 1)
        If cboPajak(Index).Text <> cboPajak(Index).List(0) And cboPajak(Index).Text <> cboPajak(Index).List(1) And cboPajak(Index).Text <> cboPajak(Index).List(2) Then
            cboPajak(Index).Text = cboPajak(Index).List(1)
        End If

Case 13 To 17, 19
    cboPajak(Index).Text = cboPajak(Index).List(cboPajak(Index).Text - 1)
    If cboPajak(Index).Text <> cboPajak(Index).List(0) And cboPajak(Index).Text <> cboPajak(Index).List(1) Then
        cboPajak(Index).Text = cboPajak(Index).List(1)
    End If
End Select
End Sub

Private Sub chEdit_Click()
On Error GoTo Salah
xJPB = Left(Trim(cboPajak(0).Text), 2)
FF1.Visible = False
FF2.Visible = False
For i = 0 To 9
    fDT(i).Visible = False
    fDT(i).Top = 120
Next
If xJPB = "03" Or xJPB = "08" Then
    FF1.Visible = True
    FF1.Top = 100
ElseIf xJPB = "02" Or xJPB = "09" Then
    FF2.Visible = True
    fDT(0).Visible = True
ElseIf xJPB = "04" Then
FF2.Visible = True
    fDT(1).Visible = True
ElseIf xJPB = "05" Then
FF2.Visible = True
    fDT(2).Visible = True
ElseIf xJPB = "06" Then
FF2.Visible = True
    fDT(3).Visible = True
ElseIf xJPB = "07" Then
FF2.Visible = True
    fDT(4).Visible = True
ElseIf xJPB = "12" Then
FF2.Visible = True
    fDT(5).Visible = True
ElseIf xJPB = "13" Then
FF2.Visible = True
    fDT(6).Visible = True
ElseIf xJPB = "15" Then
FF2.Visible = True
    fDT(7).Visible = True
ElseIf xJPB = "16" Then
FF2.Visible = True
    fDT(8).Visible = True
ElseIf xJPB = "17" Then
FF2.Visible = True
    fDT(9).Visible = True
End If
If (xJPB = 1 Or xJPB = 10 Or xJPB = 11 Or xJPB = 14) Or ((xJPB = 2 Or xJPB = 4 Or xJPB = 5 Or xJPB = 7 Or xJPB = 9) And cStandard.Value = 0 And txtPajak(4).Text <= 4) Then Exit Sub
If chEdit.Value = 1 Then
    Me.Width = 18600
Else
    Me.Width = 10395
End If
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub chPajak_Click(Index As Integer)
On Error GoTo Salah
Select Case Index
Case 0
    
    If chPajak(0).Value = 1 Then
        chPajak(1).Value = 0
        chPajak(2).Value = 0
        
        If chPajak(1).Enabled = True Then
            CTANYA = MsgBox("Memasukkan data bangunan ?" & _
               vbCrLf & "Yes : Entri data bangunan atas objek pajak baru(1)" & _
               vbCrLf & "No : Entri data bangunan atas objek pajak yang lama(0)", vbYesNoCancel + vbQuestion, "Entri Data")
            If CTANYA = vbYes Then
                xxLanjut = 1
            Else
                xxLanjut = 0
            End If
        End If
    Else
        If chPajak(1).Value = 0 And chPajak(2).Value = 0 Then
            chPajak(0).Value = 1
        End If
    End If
Case 1
    If chPajak(1).Value = 1 Then
        chPajak(0).Value = 0
        chPajak(2).Value = 0
        
'        xTanya = MsgBox("Apa anda yakin melakukan pemutakhiran OBJEK PAJAK?", vbQuestion + vbYesNo, "Penghapusan NOP")
'        If xTanya = vbYes Then
'            xID = 2
'           ' frmList_Objek.Show
'        Else
'            chPajak(0).Value = 1
'            chPajak(1).Value = 0
'        End If
    Else
        If chPajak(0).Value = 0 And chPajak(2).Value = 0 Then
            chPajak(1).Value = 1
        End If
    End If
Case 2
    
   If chPajak(2).Value = 1 Then
    chPajak(0).Value = 0
    chPajak(1).Value = 0
    cmdSave.Caption = "&Hapus"
    Else
        cmdSave.Caption = "&Proses"
        If chPajak(0).Value = 0 And chPajak(2).Value = 0 Then
            chPajak(1).Value = 1
        End If
'    xTanya = MsgBox("Apa anda yakin menghapus OBJEK PAJAK?", vbQuestion + vbYesNo, "Penghapusan NOP")
'    If xTanya = vbYes Then
'        xID = 2
'       ' frmList_Objek.Show
'    Else
'        chPajak(0).Value = 1
'        chPajak(2).Value = 0
'    End If
    
   End If
Case 3
    If chPajak(3).Value = 1 Then
        txtPajak(6).Enabled = True
        txtPajak(6).BackColor = vbWhite
        
    Else
        txtPajak(6).Enabled = False
        txtPajak(6).BackColor = vbButtonFace
    End If
End Select

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
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = 0
    End If
Next
For i = 7 To 9
    txtPajak(i).Text = "-"
Next
tKet.Text = "-"
txtPajak(4).Text = 1
xDinding = 0: xLantai = 0: xAtap = 0: xLangit2 = 0
umur_EFF = 0
FListrik = 0: FAC1 = 0: FAC2 = 0: FAC3 = 0: nDBKB = 0: xSUSUT = 0
DAYA_LISTRIK = 0: JUM_SPLIT = 0: JUM_WINDOW = 0
LUAS_HRINGAN = 0: LUAS_HSEDANG = 0: LUAS_HBERAT = 0: LUAS_HPENUTUP = 0
JUM_LAP_BETON1 = 0: JUM_LAP_BETON2 = 0: JUM_LAP_ASPAL1 = 0: JUM_LAP_ASPAL2 = 0: JUM_LAP_RUMPUT1 = 0: JUM_LAP_RUMPUT2 = 0: JUM_LAP_BETON11 = 0: JUM_LAP_BETON21 = 0: JUM_LAP_ASPAL11 = 0: JUM_LAP_ASPAL21 = 0: JUM_LAP_RUMPUT11 = 0: JUM_LAP_RUMPUT21 = 0
PANJANG_PAGAR = 0: BAHAN_PAGAR1 = 0: BAHAN_PAGAR2 = 0: LEBAR_TANGGA1 = 0: LEBAR_TANGGA2 = 0: JUM_LIFT1 = 0: JUM_LIFT2 = 0: JUM_LIFT3 = 0: JUM_PABX = 0: DALAM_SUMUR = 0
BAKAR_H = 0: BAKAR_S = 0: BAKAR_F = 0
nSistem = 0: Luas_Kolam = 0: JUM_AC_CENTRAL = 0: JUM_GENSET = 0: Nil_Boiler_Ht = 0: Nil_Boiler_Ap = 0: nMezanin = 0: nDUKUNG = 0
totChar = 0
cStandard.Value = 0
cboPajak(0).Text = cboPajak(0).List(0)
For i = 13 To 17
    cboPajak(i).Text = cboPajak(i).List(1)
Next
cboPajak(19).Text = cboPajak(19).List(1)
cTPajak.Text = cTPajak.List(0)
cboPajak(2).Text = cboPajak(2).List(0)
cboPajak(3).Text = 0
cboPajak(4).Text = cboPajak(4).List(1)
cboPajak(5).Text = cboPajak(5).List(1)
cboPajak(6).Text = cboPajak(6).List(4)
cboPajak(7).Text = cboPajak(7).List(1)
cboPajak(8).Text = cboPajak(8).List(4)
cboPajak(9).Text = cboPajak(9).List(1)

cJPB29.Clear
cJPB29.Text = "1"
cJPB4.Clear
cJPB4.Text = "1"
cJPB5.Clear
cJPB5.Text = "1"
cJPB6.Clear
cJPB6.Text = "1"
cJPB7a.Clear
cJPB7a.Text = "1"
cJPB7b.Clear
cJPB7b.Text = "0"
cJPB12.Clear
cJPB12.Text = "1"
cJPB13.Clear
cJPB13.Text = "1"
cJPB15.Clear
cJPB15.Text = "1"
cJPB16.Clear
cJPB16.Text = "1"
'NO_FORM = 0
' chPajak(1).Enabled = True
'    chPajak(2).Enabled = True
'chPajak(1).Value = 1
'chPajak(3).Value = 0
For i = 0 To 2
        dtPajak(i).Value = Format(Now, "dd/mm/yyyy")
    Next
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
Unload Me
Unload frmDBKB
xID = ""
'NO_FORM = 0
byPass = "03"
End Sub

Private Sub cmdNOP1_Click()
frmNOP.Show
End Sub

Private Sub cmdNOP2_Click()
frmNOP.Show
End Sub

Private Sub cmdNOP3_Click()
frmNOP.Show
End Sub


Private Sub cmdHitung_Click()
On Error GoTo Salah
'tListrik.Text = "@Rp. " & DAYA_LISTRIK
'Hitung Nilai Daya Listrik
Dim nMaterial, nFAS1, nFAS2, nSusut, nSusut1
Dim W0, W1, W2, W3, W4, W5, W6, W7, W8, W9
Dim WW0, WW1, WW2, WW3, WW4, WW5, WW6, WW7, WW8, WW9
Dim s_W0, s_W4, s_W5, s_W6, s_W7, s_W8, s_W9, s_WW0
Dim QQ1, QQ2, QQ3, QQ4, QQ5, QQ6
Dim s_QQ1, s_QQ2, s_QQ3, s_QQ4, s_QQ5, s_QQ6
Dim n_Mezanin, s_nFas2
cmdProses_Click
DBKB_FAS1
DBKB_FAS2
DBKB_FAS3
W1 = DAYA_LISTRIK * tListrik.Text * 1 / 1000
'Hitung Nilai AC SPlit
W2 = JUM_SPLIT * tAC1.Text
'Hitung Nilai AC WINDOW
W3 = JUM_SPLIT * tAC2.Text
'Hitung Nilai Pekerasan Halaman
W4 = (LUAS_HRINGAN * tHal1.Text) + (LUAS_HSEDANG * tHal2.Text) + (LUAS_HBERAT * tHal3.Text) + (LUAS_HPENUTUP * tHal4.Text)
If tHal1.Text = 0 Or tHal1.Text = "" Then LUAS_HRINGAN = 0
If tHal2.Text = 0 Or tHal2.Text = "" Then LUAS_HSEDANG = 0
If tHal3.Text = 0 Or tHal3.Text = "" Then LUAS_HBERAT = 0
If tHal4.Text = 0 Or tHal4.Text = "" Then LUAS_HPENUTUP = 0
s_W4 = LUAS_HRINGAN + LUAS_HSEDANG + LUAS_HBERAT + LUAS_HPENUTUP   'Susut/M
'Hitung Nilai Pagar
If cboPajak(16).Text = cboPajak(16).List(0) Then
    W5 = BAHAN_PAGAR1 * tPagar.Text
    If tPagar.Text = 0 Or tPagar.Text = "" Then BAHAN_PAGAR1 = 0
    s_W5 = BAHAN_PAGAR1
Else
    W5 = BAHAN_PAGAR2 * tPagar.Text
    If tPagar.Text = 0 Or tPagar.Text = "" Then BAHAN_PAGAR2 = 0
    s_W5 = BAHAN_PAGAR2
End If
'Hitung Nilai Tangga Berjalan
    W6 = (LEBAR_TANGGA1 * tTangga1.Text) + (LEBAR_TANGGA2 * tTangga2.Text)
    If tTangga1.Text = 0 Or tTangga1.Text = "" Then LEBAR_TANGGA1 = 0
    If tTangga2.Text = 0 Or tTangga2.Text = "" Then LEBAR_TANGGA2 = 0
    W7 = (JUM_PABX * tPABX.Text) + (DALAM_SUMUR * tSumur.Text)
    If tPABX.Text = 0 Or tPABX.Text = "" Then JUM_PABX = 0
    If tSumur.Text = 0 Or tSumur.Text = "" Then DALAM_SUMUR = 0
    s_W6 = LEBAR_TANGGA1 + LEBAR_TANGGA2
    s_W7 = JUM_PABX + DALAM_SUMUR
'Hitung Nilai Proteksi API
    If Left(Trim(cboPajak(13).Text), 2) = "01" Then 'cboPajak(13).List(0) Then
        HW1 = BAKAR_H 'txtPajak(3).Text * BAKAR_H
    Else
        HW1 = 0
    End If
    If Left(cboPajak(14).Text, 2) = "01" Then 'cboPajak(14).List(0) Then
        HW2 = BAKAR_S 'txtPajak(3).Text * BAKAR_S
    Else
        HW2 = 0
    End If
    If Left(cboPajak(15).Text, 2) = "01" Then 'cboPajak(15).List(0) Then
        HW3 = BAKAR_F 'txtPajak(3).Text * BAKAR_F
    Else
        HW3 = 0
    End If
W8 = HW1 + HW2 + HW3
s_W8 = W8
    
'Hitung Nilai Lapangan Tenis
If tLap_Beton1.Text * 1 > 1 Then
    QQ1 = tLap_Beton1.Text * JUM_LAP_BETON11
    If tLap_Beton1.Text = 0 Or tLap_Beton1.Text = "" Then JUM_LAP_BETON11 = 0
    s_QQ1 = JUM_LAP_BETON11
Else
    QQ1 = tLap_Beton1.Text * JUM_LAP_BETON1
    If tLap_Beton1.Text = 0 Or tLap_Beton1.Text = "" Then JUM_LAP_BETON1 = 0
    s_QQ1 = JUM_LAP_BETON1
End If

If tLap_Beton2.Text * 1 > 1 Then
    QQ2 = tLap_Beton2.Text * JUM_LAP_BETON21
    If tLap_Beton2.Text = 0 Or tLap_Beton2.Text = "" Then JUM_LAP_BETON21 = 0
    s_QQ2 = JUM_LAP_BETON21
Else
    QQ2 = tLap_Beton2.Text * JUM_LAP_BETON2
    If tLap_Beton2.Text = 0 Or tLap_Beton2.Text = "" Then JUM_LAP_BETON2 = 0
    s_QQ2 = JUM_LAP_BETON2
End If

If tLap_Aspal1.Text * 1 > 1 Then
    QQ3 = tLap_Aspal1.Text * JUM_LAP_ASPAL11
    If tLap_Aspal1.Text = 0 Or tLap_Aspal1.Text = "" Then JUM_LAP_ASPAL11 = 0
    s_QQ3 = JUM_LAP_ASPAL11
Else
    QQ3 = tLap_Aspal1.Text * JUM_LAP_ASPAL1
    If tLap_Aspal1.Text = 0 Or tLap_Aspal1.Text = "" Then JUM_LAP_ASPAL1 = 0
    s_QQ3 = JUM_LAP_ASPAL1
End If
If tLap_Aspal2.Text * 1 > 1 Then
    QQ4 = tLap_Aspal2.Text * JUM_LAP_ASPAL21
    If tLap_Aspal2.Text = 0 Or tLap_Aspal2.Text = "" Then JUM_LAP_ASPAL21 = 0
    s_QQ4 = JUM_LAP_ASPAL21
Else
    QQ4 = tLap_Aspal2.Text * JUM_LAP_ASPAL2
    If tLap_Aspal2.Text = 0 Or tLap_Aspal2.Text = "" Then JUM_LAP_ASPAL2 = 0
    s_QQ4 = JUM_LAP_ASPAL2
End If
If tLap_Tanah1.Text * 1 > 1 Then
    QQ5 = tLap_Tanah1.Text * JUM_LAP_RUMPUT11
    If tLap_Tanah1.Text = 0 Or tLap_Tanah1.Text = "" Then JUM_LAP_RUMPUT11 = 0
    s_QQ5 = JUM_LAP_RUMPUT11
Else
    QQ5 = tLap_Tanah1.Text * JUM_LAP_RUMPUT1
    If tLap_Tanah1.Text = 0 Or tLap_Tanah1.Text = "" Then JUM_LAP_RUMPUT1 = 0
    s_QQ5 = JUM_LAP_RUMPUT1
End If
If tLap_Tanah2.Text * 1 > 1 Then
    QQ6 = tLap_Tanah2.Text * JUM_LAP_RUMPUT21
    If tLap_Tanah2.Text = 0 Or tLap_Tanah2.Text = "" Then JUM_LAP_RUMPUT21 = 0
    s_QQ6 = JUM_LAP_RUMPUT21
Else
    QQ6 = tLap_Tanah2.Text * JUM_LAP_RUMPUT2
    If tLap_Tanah2.Text = 0 Or tLap_Tanah2.Text = "" Then JUM_LAP_RUMPUT2 = 0
    s_QQ6 = JUM_LAP_RUMPUT2
End If
W9 = QQ1 + QQ2 + QQ3 + QQ4 + QQ5 + QQ6
s_W9 = s_QQ1 + s_QQ2 + s_QQ3 + s_QQ4 + s_QQ5 + s_QQ6
'Hitung Nilai Lift

'If txtPajak(4).Text * 1 < 5 Then
'    AA1 = tLift1.Text * JLIFT(1)
'    AA2 = tLift2.Text * JLIFT(5)
'    AA3 = tLift3.Text * JLIFT(9)
'ElseIf txtPajak(4).Text * 1 >= 5 And txtPajak(4).Text * 1 <= 9 Then
'    AA1 = tLift1.Text * JLIFT(2)
'    AA2 = tLift2.Text * JLIFT(6)
'    AA3 = tLift3.Text * JLIFT(10)
'ElseIf txtPajak(4).Text * 1 >= 10 And txtPajak(4).Text * 1 <= 19 Then
'    AA1 = tLift1.Text * JLIFT(3)
'    AA2 = tLift2.Text * JLIFT(7)
'    AA3 = tLift3.Text * JLIFT(11)
'Else
'    AA1 = tLift1.Text * JLIFT(4)
'    AA2 = tLift2.Text * JLIFT(8)
'    AA3 = tLift3.Text * JLIFT(12)
'End If
'w0=AA1+AA2+AA3
W0 = (JLIFT(1) * tLift1.Text) + (JLIFT(2) * tLift2.Text) + (JLIFT(3) * tLift3.Text)
If tLift1.Text = 0 Or tLift1.Text = "" Then JLIFT(1) = 0
If tLift2.Text = 0 Or tLift2.Text = "" Then JLIFT(2) = 0
If tLift3.Text = 0 Or tLift3.Text = "" Then JLIFT(3) = 0
s_W0 = JLIFT(1) + JLIFT(2) + JLIFT(3)
'If tLift1.Text = 0 Then JLIFT(1) = 0
'If tLift2.Text = 0 Then JLIFT(2) = 0
'If tLift3.Text = 0 Then JLIFT(3) = 0
'w0 = (JLIFT(1)) + (JLIFT(2)) + (JLIFT(3)) 'DBKB Berdasarkan Range, Tidak Dikalikan Jumlah Liff
'        If tKolam.Text < 51 Then
'            bb1 = tKolam.Text * Luas_Kolam(6)
'        ElseIf tKolam.Text >= 51 And tKolam.Text <= 100 Then
'            bb1 = tKolam.Text * Luas_Kolam(7)
'        ElseIf tKolam.Text >= 101 And tKolam.Text <= 200 Then
'            bb1 = tKolam.Text * Luas_Kolam(8)
'        ElseIf tKolam.Text >= 201 And tKolam.Text <= 400 Then
'            bb1 = tKolam.Text * Luas_Kolam(9)
'        Else
'            bb1 = tKolam.Text * Luas_Kolam(10)
'        End If
'End If
'If Left(cboPajak(19).Text, 2) = "01" Then
'If tKolam.Text = 0 Then Luas_Kolam = 0

'Kolam Renang
    WW0 = Luas_Kolam * tKolam.Text
    If tKolam.Text = 0 Or tKolam.Text = "" Then Luas_Kolam = 0
    s_WW0 = Luas_Kolam
'Cek AC Central Untuk Bangunan Standard
    If Left(cboPajak(17).Text, 2) = "01" Or cboPajak(17).ListIndex = 0 Then
        WW1 = JUM_AC_CENTRAL
    Else
        WW1 = 0
    End If
    If tGenset.Text = 0 Then JUM_GENSET = 0
    'Genset di Bangunan
    WW2 = JUM_GENSET '* tGenset.Text
    
    'Boiler Hotel dan Apartemen
    If xJPB = 7 Then
        WW3 = Nil_Boiler_Ht * JPB7d.Text 'tBoiler.Text
    ElseIf xJPB = 13 Then
        WW3 = Nil_Boiler_Ap * JPB13d.Text
    Else
        WW3 = 0
    End If
'AC Central Untuk Bangunan Non Standard
If (txtPajak(4).Text * 1 > 4 Or cStandard.Value = 1) Then
If Left(cboPajak(17).Text, 2) = "01" Or cboPajak(17).ListIndex = 0 Then
    WW1 = 0
    If xJPB = 2 Then 'Perkantoran
        WW4 = Nil_AC_Central(1) '* txtPajak(3).Text
    ElseIf xJPB = 4 Then 'Pertokoan
        WW4 = Nil_AC_Central(4) '* txtPajak(3).Text
    ElseIf xJPB = 5 Then 'Rumah Sakit
        WW4 = Nil_AC_Central(5)
        WW5 = Nil_AC_Central(6)
    ElseIf xJPB = 7 Then 'Hotel
        WW4 = Nil_AC_Central(2)
        WW5 = Nil_AC_Central(3)
    ElseIf xJPB = 13 Then 'Apartemen
        WW4 = Nil_AC_Central(7)
        WW5 = Nil_AC_Central(8)
    End If
Else
    WW4 = 0
    WW5 = 0
End If
End If
'Nilai Tambahan

If xJPB = 3 Or xJPB = 8 Then
    nDUKUNG = nDUKUNG ' * (xSUSUT / 100)
    nMezanin = nMezanin ' * (xSUSUT / 100)
Else
    nDUKUNG = 0
    nMezanin = 0
End If

nMaterial = (xAtap * 1) + (xDinding * 1) + (xLantai * 1) + (xLangit2 * 1)
'nFAS = W0 + W1 + W2 + W3 + W4 + W5 + W6 + w7 + W8 + W9 + WW0 + WW1 + WW2 + WW3 + WW4 + WW5
nFAS1 = W1 + W2 + W3 + WW1 + WW2 + WW3 + WW4 + WW5 'Fasilitas Tidak Disusutkan
nFAS2 = W0 + W4 + W5 + W6 + W7 + W8 + W9 + WW0 'Fasilitas Disusutkan
s_nFas2 = s_W0 + s_W4 + s_W5 + s_W6 + s_W7 + s_W8 + s_W9 + s_WW0 'Fasilitas Disusutkan
'If txtPajak(4).Text > 4 Or cStandard.Value = 1 Or XJPB = 14 Or XJPB = 15 Then
'    nMaterial = 0
'End If
If Left(cboPajak(5).Text, 2) * 1 = 4 And ck_Ulin = 0 Then nDBKB = nDBKB * 0.7
JGuna = Left(Trim(cboPajak(0).Text), 2) * 1
JLANTAI = txtPajak(4).Text * 1
'Menentukan Pengganti Baru Bangunan dan Nilai Penyusutan
'-------------------------------------
'Bangunan Standar
'-------------------------------------
If txtPajak(3).Text = "" Or txtPajak(3).Text = 0 Then txtPajak(3).Text = 1
If JPB38(4).Text = 0 Then n_Mezanin = 0 Else n_Mezanin = nMezanin / JPB38(4).Text
If (JGuna = 1 Or JGuna = 3 Or JGuna = 8 Or JGuna = 10 Or JGuna = 11 Or JGuna = 2 Or JGuna = 4 Or JGuna = 5 Or JGuna = 7 Or JGuna = 9) And JLANTAI <= 4 And (xJPB <> 14 Or xJPB <> 15) Then
    N_BANGUNAN = ((nDBKB * 1 + nMaterial * 1)) + (nDUKUNG * 1 / txtPajak(3).Text) + (n_Mezanin * 1 + s_nFas2 * 1)
Else
    nMaterial = 0
    N_BANGUNAN = ((nDBKB * 1)) + (nDUKUNG * 1 / txtPajak(3).Text) + (n_Mezanin * 1 + s_nFas2 * 1)
End If
    tampil_Susut (N_BANGUNAN * 1000)
If JGuna = 15 And xSUSUT > 50 Then
    xSUSUT = 50
End If
nSusut1 = xSUSUT / 100 * nFAS2 '(W0 + W4 + W5 + W6 + w7 + W8 + W9 + WW0 + WW2 + WW3)
'Bangunan Non Standard
nBangunan = ((nDBKB * 1 + nMaterial) * txtPajak(3).Text) + nDUKUNG + nMezanin

nSusut = nSusut1 + (xSUSUT / 100 * nBangunan)
'MsgBox nDBKB & ":" & nMaterial & ":" & txtPajak(3).Text & ":" & nDUKUNG & ":" & nMezanin & " suSUT:" & nSusut1
nSistem = (nFAS1 + nFAS2 + nBangunan) - nSusut
If xJPB = 17 Then nSistem = nDBKB
If xJPB = 11 Then nSistem = 0
'If Left(cboPajak(5).Text, 2) * 1 = 4 Then nSistem = nSistem * 0.7
'List1.AddItem w0 & ":" & w1 & ":" & w2 & ":" & w3 & ":" & w4 & ":" & w5 & ":" & w6 & ":" & w7 & ":" & w8 & ":" & w9 & ":" & bb1
LFas.Caption = Format(nSistem * 1000, "#,#0.00")
LSusut.Caption = Format(nSusut, "#,#0.00")
txtPajak(5).Text = Format(nSistem, "#,#0.00")
'MsgBox " BETON 1 : " & JUM_LAP_BETON1 & _
       vbCrLf & " BETON 2 : " & JUM_LAP_BETON2 & _
       vbCrLf & " ASPAL 1 : " & JUM_LAP_ASPAL1 & _
       vbCrLf & " ASPAL 2 : " & JUM_LAP_ASPAL2 & _
       vbCrLf & " RUMPUT 1 : " & JUM_LAP_RUMPUT1 & _
       vbCrLf & " RUMPUT 2 : " & JUM_LAP_RUMPUT2 & _
       vbCrLf & " BETON 11 : " & JUM_LAP_BETON11 & _
       vbCrLf & " BETON 21 : " & JUM_LAP_BETON21 & _
       vbCrLf & " ASPAL 11 : " & JUM_LAP_ASPAL11 & _
       vbCrLf & " ASPAL 21 : " & JUM_LAP_ASPAL21 & _
       vbCrLf & " RUMPUT 11 : " & JUM_LAP_RUMPUT11 & _
       vbCrLf & " RUMPUT 21 : " & JUM_LAP_RUMPUT21

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub cmdNOP_Click()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
J_Karakter
If Len(Trim(txtPajak(1).Text)) - (totChar * 1) = 24 Then

'    'StrQ = "Select * From QOBJEKPAJAK WHERE NOPQ='" & Trim(txtPajak(1).Text) & "' and ((JNS_BUMI='1' OR JNS_BUMI='4') AND TOTAL_LUAS_BNG>0)order by NOPQ asc"
'    'StrQ = "Select * From QOBJEKPAJAK WHERE NOPQ='" & Trim(txtPajak(1).Text) & "' and (JNS_BUMI='1' OR JNS_BUMI='4')order by NOPQ asc"
'    'StrQ = "Select * From QOBJEKPAJAK WHERE NOPQ='" & Trim(aNOP.Text) & "' and (JNS_BUMI='1' OR JNS_BUMI='4')order by NOPQ asc"
'    'openDB (StrQ)
'    txtPajak(1).Text = aNOP.Text
'    StrQ1 = "Select * From DAT_OP_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' ORDER BY NO_BNG*1 DESC"
'    openDB (StrQ1)
'
'    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    If rPajak.EOF Then
'        MsgBox "Data Tidak Ditemukan, Silahkan Ganti...!", vbCritical, "Info..."
'        aNOP.SetFocus
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
''    Do While Not rPajak.EOF
''        txtPajak(2).Text = rPajak!SUBJEK_PAJAK_ID '" NAMA" & vbTab & ": " & rPajak!Nm_wp & vbCrLf & " LOKASI" & vbTab & ": " & rPajak!JALAN_OP & " Blok: " & rPajak!KD_BLOK & ", RT/RW: " & rPajak!RT_OP & "/" & rPajak!RW_OP & ", " & rPajak!NM_KELURAHAN & ", KEC. " & rPajak!NM_KECAMATAN 'NAMA dan Alamat
''        rPajak.MoveNext
''    Loop
'
    cboPajak(1).Text = 1
    StrQ1 = "Select * From DAT_OP_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
    openDB (StrQ1)

    call_data
    If chPajak(1).Value = 1 Or chPajak(2).Value = 1 Or xxLanjut <> 1 Then tempLog1

Else
    xID = 3
frmLIST_Objek1.Show
End If

Screen.MousePointer = vbDefault
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub cmdProses_Click1()
On Error GoTo Salah
Dim X(100)
Dim TPajak, TRenovasi, TBangun, JLANTAI, JGuna, Umur
LTipe.Caption = ""
StrQ = "Select * From TIPE_BANGUNAN order by TIPE_BNG asc"
openDB (StrQ)
i = 0
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'JPB = Left(Trim(cboPajak(0).Text), 2) * 1
Do While Not rPajak.EOF
 If (txtPajak(3).Text * 1 >= rPajak!LUAS_MIN_TIPE_BNG And txtPajak(3).Text * 1 <= rPajak!LUAS_MAX_TIPE_BNG) Then 'And txtPajak(3).Text >= rPajak!FAKTOR_PEMBAGI_TIPE_BNG Then
    i = i + 1
    'LTipe.Caption = rPajak!FAKTOR_PEMBAGI_TIPE_BNG
    X(i) = rPajak!FAKTOR_PEMBAGI_TIPE_BNG
    'Exit Sub
'ElseIf rPajak!FAKTOR_PEMBAGI_TIPE_BNG >= 549 Then
 '   LTipe.Caption = 555
  '  Exit Sub
    xN = i
    xMAX = rPajak!LUAS_MAX_TIPE_BNG
 End If
rPajak.MoveNext
Loop
For i = 1 To xN - 1
    If txtPajak(3).Text * 1 < X(i + 1) Then
        cTipe = X(i)
        GoTo cetak
    'Else
     '   cTipe = x(xN)
    End If
Next
cetak:
If txtPajak(3).Text * 1 >= X(xN) Then
    cTipe = X(xN)
End If
LTipe.Caption = cTipe

callDBKBUtama
Call CALL_DINDING("21", cboPajak(7).Text)
Call CALL_LANTAI("22", cboPajak(8).Text)
Call CALL_ATAP("23", cboPajak(6).Text)
Call CALL_LANGIT2("24", cboPajak(9).Text)
'FListrik = tListrik.Text / 1000 * Biaya
TPajak = cTPajak.Text * 1
TBangun = cboPajak(2).Text * 1
TRenovasi = cboPajak(3).Text * 1
JGuna = Left(cboPajak(0).Text, 2) * 1
JLANTAI = txtPajak(4).Text * 1
'MsgBox JLantai & ":" & JGuna & ":" & TPajak - TBangun & ":" & TRenovasi - TBangun
If TRenovasi <= 0 Then 'Tidak Ada Renovasi
    If (JGuna = 1 Or JGuna = 3 Or JGuna = 8 Or JGuna = 10 Or JGuna = 11) And JLANTAI <= 4 Then
        Umur = TPajak - TBangun
    Else
        If (TPajak - TBangun) <= 10 Then
            Umur = TPajak - TBangun
        Else
            Umur = (TPajak - TBangun + 20) / 3
        End If
    End If
Else 'Ada Renovasi
    If (JGuna = 1 Or JGuna = 3 Or JGuna = 8 Or JGuna = 10 Or JGuna = 11) And JLANTAI <= 4 Then
        Umur = TPajak - TRenovasi
    Else
        If (TRenovasi - TBangun) <= 10 Then
            Umur = ((TPajak - TBangun) + (2 * (TPajak - TBangun))) / 3
        Else
            Umur = (TPajak - TBangun + 20) / 3
        End If
    End If
End If
'If txtPajak(3).Text > 0 Then 'Ada Renovasi
'
'    If (JPB * 1 = 1 Or JPB * 1 = 2 Or JPB * 1 = 10 Or JPB * 1 = 11) And txtPajak(4).Text <= 4 Then
'    'Bangunan Standar
'        U_Efektif = cTPajak.Text * 1 - cboPajak(3).Text * 1
'
'    Else
'        'Bangunan Non Standar
'        If ((txtPajak(3).Text * 1) - (txtPajak(2).Text)) > 10 Then
'            U_Efektif = (((cTPajak.Text * 1) - (cboPajak(2).Text * 1)) + (2 * 10)) / 3
'        Else
'            U_Efektif = (((cTPajak.Text * 1) - (cboPajak(2).Text * 1)) + (2 * ((cTPajak.Text * 1) - (cboPajak(3).Text * 1)))) / 3
'        End If
'
'    End If
'
'Else 'Tidak ada Renovasi
'    If (JPB * 1 <> 1 Or JPB * 1 <> 2 Or JPB * 1 <> 10 Or JPB * 1 <> 11) And txtPajak(4).Text <= 4 Then
'    'Bangunan Standar
'        U_Efektif = cTPajak.Text * 1 - cboPajak(2).Text * 1
'    Else
'        'Bangunan Non Standar
'        If ((cTPajak.Text * 1) - (txtPajak(2).Text)) > 10 Then
'            U_Efektif = (((cTPajak.Text * 1) - (cboPajak(2).Text * 1)) + (2 * 10)) / 3
'        Else
'            U_Efektif = ((cTPajak.Text * 1) - (cboPajak(2).Text * 1))
'        End If
'
'    End If
'
'End If
'
If Umur > 40 Then
    Umur = 40
End If
umur_EFF = Round(Umur, 2) ' + 0.4)
LEff.Caption = umur_EFF 'Umur & " dan " & Round(Umur + 0.4)
'tampil_Susut

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub cmdProses_Click()
On Error GoTo Salah
Dim X(100)
Dim TPajak, TRenovasi, TBangun, JLANTAI, JGuna, Umur, JL

JPB = Left(Trim(cboPajak(0).Text), 2)
Select Case JPB
Case "01", "02", "04", "05", "07", "09", "10", "11"
If txtPajak(4).Text <= 1 And cStandard.Value = 0 Then
    JL = 1
ElseIf txtPajak(4).Text <= 4 And cStandard.Value = 0 Then
    JL = 2
Else 'If txtPajak(4).Text > 4 Or cStandard.Value = 1 Then
    JL = 3
End If
If (JPB = "01" Or JPB = "10" Or JPB = "11") And JL <= 2 Then JPB = "01"
If JPB = "05" And JL <= 2 Then JPB = "05"
If (JPB = "02" Or JPB = "04" Or JPB = "07" Or JPB = "09") And JL <= 2 Then
    JPB = "02"
End If
LTipe.Caption = ""
If JL <= 2 Then
    StrQ = "SELECT DBKB_STANDARD.THN_DBKB_STANDARD, DBKB_STANDARD.KD_JPB, TIPE_BANGUNAN.TIPE_BNG, TIPE_BANGUNAN.NM_TIPE_BNG, TIPE_BANGUNAN.LUAS_MIN_TIPE_BNG, TIPE_BANGUNAN.LUAS_MAX_TIPE_BNG, TIPE_BANGUNAN.FAKTOR_PEMBAGI_TIPE_BNG, DBKB_STANDARD.KD_BNG_LANTAI, DBKB_STANDARD.NILAI_DBKB_STANDARD FROM DBKB_STANDARD INNER JOIN TIPE_BANGUNAN ON DBKB_STANDARD.TIPE_BNG = TIPE_BANGUNAN.TIPE_BNG WHERE (((DBKB_STANDARD.THN_DBKB_STANDARD)='" & cTPajak.Text * 1 & "')) AND DBKB_STANDARD.KD_JPB='" & JPB & "'"
    openDB (StrQ)
    i = 0
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        'If (txtPajak(3).Text * 1 >= rPajak!LUAS_MIN_TIPE_BNG And txtPajak(3).Text * 1 <= rPajak!LUAS_MAX_TIPE_BNG) And Mid(Trim(rPajak!KD_BNG_LANTAI), 3, 1) * 1 = JL Then '"& JPB & "_" & &"' txtPajak(4).Text * 1 <= 1 Then
        If (txtPajak(3).Text * 1 >= rPajak!LUAS_MIN_TIPE_BNG And txtPajak(3).Text * 1 <= rPajak!LUAS_MAX_TIPE_BNG) And Mid(Trim(rPajak!KD_BNG_LANTAI), 3, 1) * 1 = txtPajak(4).Text * 1 Then  '"& JPB & "_" & &"' txtPajak(4).Text * 1 <= 1 Then
           nDBKB = rPajak!NILAI_DBKB_STANDARD
           cTipe = rPajak!TIPE_BNG
        End If
     
    rPajak.MoveNext
    Loop

Else ' Jumlah Lantai Diatas 4 atau bangunan non standard
    If JPB = 2 Then ' Perkantoran Swasta
     StrQ = "SELECT * FROM DBKB_JPB2 WHERE THN_DBKB_JPB2='" & cTPajak.Text * 1 & "' AND KLS_DBKB_JPB2 ='" & cJPB29.Text & "' ORDER BY LANTAI_MIN_JPB2,LANTAI_MAX_JPB2 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
                
                CC = rPajak!KLS_DBKB_JPB2
                If (txtPajak(4).Text * 1 >= rPajak!LANTAI_MIN_JPB2 And txtPajak(4).Text * 1 <= rPajak!LANTAI_MAX_JPB2) Then
                    i = i + 1
                    'If cc * 1 = cJPB29.Text * 1 Then
                        nDBKB = rPajak!NILAI_DBKB_JPB2
                        'GoTo keluar
                    'End If
                End If
         
        rPajak.MoveNext
        Loop
            If i = 0 Then nDBKB = 0
            
    ElseIf JPB = 4 Then ' Pertokoan
     StrQ = "SELECT * FROM DBKB_JPB4 WHERE THN_DBKB_JPB4='" & cTPajak.Text * 1 & "' ORDER BY KLS_DBKB_JPB4 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
                CC = rPajak!KLS_DBKB_JPB4
                If (txtPajak(4).Text * 1 >= rPajak!LANTAI_MIN_JPB4 And txtPajak(4).Text * 1 <= rPajak!LANTAI_MAX_JPB4) Then
                    If CC * 1 = cJPB4.Text * 1 Then
                        i = i + 1
                        nDBKB = rPajak!NILAI_DBKB_JPB4
                    End If
                End If
         
        rPajak.MoveNext
        Loop
        If i = 0 Then nDBKB = 0
        
    ElseIf JPB = 5 Then ' Rumah Sakit/Klinik
     StrQ = "SELECT * FROM DBKB_JPB5 WHERE THN_DBKB_JPB5='" & cTPajak.Text * 1 & "' AND KLS_DBKB_JPB5 ='" & cJPB5.Text & "' ORDER BY KLS_DBKB_JPB5 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
                If (txtPajak(4).Text * 1 >= rPajak!LANTAI_MIN_JPB5 And txtPajak(4).Text * 1 <= rPajak!LANTAI_MAX_JPB5) Then
                        i = i + 1
                        nDBKB = rPajak!NILAI_DBKB_JPB5
                End If
         
        rPajak.MoveNext
        Loop
        If i = 0 Then nDBKB = 0
    ElseIf JPB = 7 Then ' Hotel
     If cJPB7b.Text = 0 Then cJPB7b = 5
     StrQ = "SELECT * FROM DBKB_JPB7 WHERE THN_DBKB_JPB7='" & cTPajak.Text * 1 & "' AND JNS_DBKB_JPB7 ='" & cJPB7a.Text & "' AND BINTANG_DBKB_JPB7 ='" & cJPB7b.Text & "' "
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
                If (txtPajak(4).Text * 1 >= rPajak!LANTAI_MIN_JPB7 And txtPajak(4).Text * 1 <= rPajak!LANTAI_MAX_JPB7) Then
                        i = i + 1
                        nDBKB = rPajak!NILAI_DBKB_JPB7
                End If
         
        rPajak.MoveNext
        Loop
        If i = 0 Then nDBKB = 0
        
    End If
   
End If
Case "03"
call_Mezz
call_Dukung


      StrQ = "SELECT * FROM DBKB_JPB3 WHERE THN_DBKB_JPB3='" & cTPajak.Text * 1 & "' " 'AND (LBR_BENT_MIN_DBKB_JPB3 *1>= '" & JPB38(1).Text * 1 & "' AND LBR_BENT_MAX_DBKB_JPB3*1<='" & JPB38(1).Text * 1 & "')"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
            If JPB38(1).Text >= rPajak!LBR_BENT_MIN_DBKB_JPB3 And JPB38(1).Text <= rPajak!LBR_BENT_MAX_DBKB_JPB3 Then
                If JPB38(0).Text * 1 >= rPajak!TING_KOLOM_MIN_DBKB_JPB3 * 1 And JPB38(0).Text * 1 <= rPajak!TING_KOLOM_MAX_DBKB_JPB3 * 1 Then
                    i = i + 1
                    nDBKB = rPajak!NILAI_DBKB_JPB3
                End If
            End If
        rPajak.MoveNext
        Loop
        If i = 0 Then nDBKB = 0
        'MsgBox i
    
Case "06"
      StrQ = "SELECT * FROM DBKB_JPB6 WHERE THN_DBKB_JPB6='" & cTPajak.Text * 1 & "' AND KLS_DBKB_JPB6 ='" & cJPB6.Text & "' ORDER BY KLS_DBKB_JPB6 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
            i = i + 1
            nDBKB = rPajak!NILAI_DBKB_JPB6
        rPajak.MoveNext
        Loop
        If i = 0 Then nDBKB = 0

Case "08"
call_Mezz
call_Dukung

      StrQ = "SELECT * FROM DBKB_JPB8 WHERE THN_DBKB_JPB8='" & cTPajak.Text * 1 & "' ORDER BY LBR_BENT_MIN_DBKB_JPB8 ASC" 'AND (LBR_BENT_MIN_DBKB_JPB8>='" & JPB38(1).Text * 1 & "' AND LBR_BENT_MAX_DBKB_JPB8<='" & JPB38(1).Text * 1 & "') ORDER BY LBR_BENT_MIN_DBKB_JPB8 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
            If JPB38(1).Text >= rPajak!LBR_BENT_MIN_DBKB_JPB8 And JPB38(1).Text <= rPajak!LBR_BENT_MAX_DBKB_JPB8 Then
                If JPB38(0).Text * 1 >= rPajak!TING_KOLOM_MIN_DBKB_JPB8 And JPB38(0).Text * 1 <= rPajak!TING_KOLOM_MAX_DBKB_JPB8 Then
                    i = i + 1
                    nDBKB = rPajak!NILAI_DBKB_JPB8
                End If
            End If
        rPajak.MoveNext
        Loop
        If i = 0 Then nDBKB = 0
        

Case "11" ' Tidak Kena Pajak
    nDBKB = 0
Case "12" 'Bangunan Parkir
      StrQ = "SELECT * FROM DBKB_JPB12 WHERE THN_DBKB_JPB12='" & cTPajak.Text * 1 & "' AND TYPE_DBKB_JPB12 ='" & cJPB12.Text & "' ORDER BY TYPE_DBKB_JPB12 ASC"
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
            i = i + 1
            nDBKB = rPajak!NILAI_DBKB_JPB12
        rPajak.MoveNext
        Loop
        If i = 0 Then nDBKB = 0
Case "13"
         StrQ = "SELECT * FROM DBKB_JPB13 WHERE THN_DBKB_JPB13='" & cTPajak.Text * 1 & "' AND KLS_DBKB_JPB13='" & cJPB13.Text & "' "
        openDB (StrQ)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
                If (txtPajak(4).Text * 1 >= rPajak!LANTAI_MIN_JPB13 And txtPajak(4).Text * 1 <= rPajak!LANTAI_MAX_JPB13) Then
                        i = i + 1
                        nDBKB = rPajak!NILAI_DBKB_JPB13
                End If
         
        rPajak.MoveNext
        Loop
        If i = 0 Then nDBKB = 0

Case "14" 'Kanopi Pompa Bensin
    StrQ = "select * from DBKB_JPB14 where THN_DBKB_JPB14='" & cTPajak.Text & "'"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        nDBKB = rPajak!NILAI_DBKB_JPB14
        rPajak.MoveNext
    Loop
Case "15" 'Tangki Minyak
    StrQ = "select * from DBKB_JPB15 where THN_DBKB_JPB15='" & cTPajak.Text & "' AND JNS_TANGKI_DBKB_JPB15='" & cJPB15.Text * 1 & "'"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
        If JPB15.Text * 1 >= rPajak!KAPASITAS_MIN_DBKB_JPB15 And JPB15.Text * 1 <= rPajak!KAPASITAS_MAX_DBKB_JPB15 Then
            i = i + 1
            nDBKB = rPajak!NILAI_DBKB_JPB15
        End If
        rPajak.MoveNext
    Loop
    If i = 0 Then nDBKB = 0
    
 Case "16" 'Gedung Sekolah
    StrQ = "select * from DBKB_JPB16 where THN_DBKB_JPB16='" & cTPajak.Text & "' AND KLS_DBKB_JPB16='" & cJPB16.Text * 1 & "'"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    i = 0
    Do While Not rPajak.EOF
        If txtPajak(4).Text * 1 >= rPajak!LANTAI_MIN_JPB16 And txtPajak(4).Text * 1 <= rPajak!LANTAI_MAX_JPB16 Then
            i = i + 1
            nDBKB = rPajak!NILAI_DBKB_JPB16
        End If
        rPajak.MoveNext
    Loop
    If i = 0 Then nDBKB = 0
   
Case "17" 'Tower/Menara Telekomunikasi
    StrQ = "select * from DBKB_JPB17 where THN_DBKB_JPB17='" & cTPajak.Text & "' AND (TINGGI_MIN_JPB17*1>='" & cJPB17.Text * 1 & "' and TINGGI_MAX_JPB17*1<='" & cJPB17.Text * 1 & "')"
    openDB (StrQ)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    'i = 0
    
    Do While Not rPajak.EOF
        'If txtPajak(4).Text * 1 >= rPajak!LANTAI_MIN_JPB16 And txtPajak(4).Text * 1 <= rPajak!LANTAI_MAX_JPB16 Then
         '   i = i + 1
            nMenara = rPajak!NILAI_BNG_MENARA_JPB17
            nMekanikal = rPajak!BIAYA_MEKANIK_JPB17
            nPagar = rPajak!NILAI_BGN_PAGAR_JPB17
            nDBKB = nMenara + nMekanikal + nPagar 'rPajak!NILAI_DBKB_JPB17
        'End If
        rPajak.MoveNext
    Loop
    

    
End Select
'MsgBox JPB & ":" & JL
LTipe.Caption = cTipe
LDBKB.Caption = nDBKB
Call CALL_DINDING("21", cboPajak(7).Text)
Call CALL_LANTAI("22", cboPajak(8).Text)
Call CALL_ATAP("23", cboPajak(6).Text)
Call CALL_LANGIT2("24", cboPajak(9).Text)
'FListrik = tListrik.Text / 1000 * Biaya
TPajak = cTPajak.Text * 1
TBangun = cboPajak(2).Text * 1
TRenovasi = cboPajak(3).Text * 1
JGuna = Left(Trim(cboPajak(0).Text), 2) * 1
JLANTAI = txtPajak(4).Text * 1
'MsgBox JLantai & ":" & JGuna & ":" & TPajak - TBangun & ":" & TRenovasi - TBangun
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
    ElseIf TBangun > 0 And TRenovasi <= 0 Then
        If TPajak - TBangun > 10 Then
            Umur = ((TPajak - TBangun) + (2 * 10)) / 3
        Else
            Umur = TPajak - TBangun
        End If
    Else
        MsgBox "Tahun Pembangunan Tidak Boleh Kosong...", vbCritical, "Tetnong..."
            cTPajak.Text = Format(Now, "yyyy")
            cTPajak.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
'If TRenovasi <= 0 Then 'Tidak Ada Renovasi
'    If (JGuna = 1 Or JGuna = 3 Or JGuna = 8 Or JGuna = 10 Or JGuna = 11 Or JGuna = 2 Or JGuna = 4 Or JGuna = 5 Or JGuna = 7 Or JGuna = 9) And JLANTAI <= 4 Then
'        Umur = TPajak - TBangun
'    Else
'        If (TPajak - TBangun) <= 10 Then
'            Umur = TPajak - TBangun
'        Else
'            Umur = (TPajak - TBangun + 20) / 3
'        End If
'    End If
'Else 'Ada Renovasi
'    If (JGuna = 1 Or JGuna = 3 Or JGuna = 8 Or JGuna = 10 Or JGuna = 11 Or JGuna = 2 Or JGuna = 4 Or JGuna = 5 Or JGuna = 7 Or JGuna = 9) And JLANTAI <= 4 Then
'        Umur = TPajak - TRenovasi
'    Else
'        If (TRenovasi - TBangun) <= 10 Then
'            Umur = ((TPajak - TBangun) + (2 * (TPajak - TBangun))) / 3
'        Else
'            Umur = (TPajak - TBangun + 20) / 3
'        End If
'    End If
'End If

If Umur > 40 Then
    Umur = 40
End If
umur_EFF = Round(Umur) ' + 0.4)
LEff.Caption = umur_EFF 'Umur & " dan " & Round(Umur + 0.4)
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description


End Sub


Private Sub cmdSave_Click()
On Error GoTo Salah
If cmdSave.Caption = "&Proses" Then
    For Each Control In Me
        If TypeOf Control Is TextBox Then
            If Control.Text = "" Then
                Control.Text = 0
            End If
        End If
        If TypeOf Control Is ComboBox Then
            If Control.Text = "" Then
                Control.Text = Control.List(0)
            End If
        End If
    Next
    StrQ = "Select * From DAT_OP_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' and (NO_BNG*1='" & cboPajak(1).Text * 1 & "')ORDER BY NO_BNG*1 DESC"
    openDB (StrQ)
    If chPajak(0).Value = 1 Then
        cTrans = 1
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        If Not rPajak.EOF Then
            MsgBox "Nomor Bangunan Sudah Digunakan, " & _
                vbCrLf & "Apabila NOP digunakan untuk lebih dari 1 Bangunan", vbCritical, "Tetnong.."
            Exit Sub
            If Left(Trim(cboPajak(0).Text), 2) <> rPajak!KD_JPB Then
                MsgBox "Anda tidak boleh menggunakan NOP yang sama untuk" & _
                    vbCrLf & "Jenis Penggunaan Bangunan (JPB) berbeda..."
                    Exit Sub
            End If
        End If
        StrQ1 = "Select * From DAT_OP_BUMI WHERE KD_PROPINSI  + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_URUT ASC"
        openDB (StrQ1)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        If rPajak.EOF Then
            MsgBox "Nomor Objerk Pajak tidak terdaftar...", vbCritical, "Tetnong...!"
            Exit Sub
        End If
        
    Else
        cTrans = 2
        If rPajak.EOF Then
            MsgBox "Data Tidak Dapat Ditemukan,,,!", vbCritical, "Info..."
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    
    cmdHitung_Click
    If txtPajak(0).Text = "" Or txtPajak(0).Text = "0" Then
        MsgBox "Nomor Formulir LSPOP Masih Kosong...!", vbCritical, "Tetnong..."
        Exit Sub
    End If
    
    J_Karakter
    If Len(Trim(txtPajak(1).Text)) - (totChar * 1) < 24 Then
    'If Or Len(txtPajak(1).Text) <> 24 Then
        MsgBox "Format Nomor Objek Pajak Belum Benar...!", vbCritical, "Tetnong..."
        Exit Sub
    ElseIf txtPajak(1).Text = "" Or txtPajak(1).Text = "0" Then
        MsgBox "Nomor Objek Pajak Masih Kosong...!", vbCritical, "Tetnong..."
        Exit Sub
    End If
'StrQ = "Select * From QOBJEKPAJAK WHERE NOPQ='" & Trim(txtPajak(1).Text) & "' and ((JNS_BUMI='1' OR JNS_BUMI='4') AND TOTAL_LUAS_BNG>0)order by NOPQ asc"
'StrQ = "Select * From QOBJEKPAJAK WHERE NOPQ='" & Trim(aNOP.Text) & "' and (JNS_BUMI='1' OR JNS_BUMI='4') order by NOPQ asc"
'    openDB (StrQ)
'    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'    If rPajak.EOF Then
'        MsgBox "NOP Tidak Ada Di Database" & _
'        vbCrLf & "Atau Jenis Penggunan Objek Tidak Untuk Bangunan...!", vbCritical, "Info..."
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    TANYA = MsgBox("Apa Anda Yakin Proses Penilaian Bangunan?", vbInformation + vbYesNo, "Saved...")
    If TANYA = vbNo Then
        Exit Sub
    End If
    
    frmDBKB.Show
    xID = 110
Else
    TANYA = MsgBox("Apa Anda Yakin Menghapus Data Bangunan Ini?", vbInformation + vbYesNo, "Deleted...")
    If TANYA = vbNo Then
        Exit Sub
    End If
    'DEL_BANGUNAN1
    'DEL_INDIVIDU1
    'DEL_FASILITAS1
    'DEL_TAMBAHAN1
   ' UP_OBJEK
    xxJPB = Left(cboPajak(0).Text, 2) * 1
    
   ' c_str = "HAPUS_BANGUNAN '" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "'," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "'," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "'," & _
    "0, 0,'" & Trim(aNOP.Text) & "'," & _
    "'" & xxJPB & "','" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "'," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "'"
    xxKec = Mid(Trim(aNOP.Text), 7, 3)
    xxKel = Mid(Trim(aNOP.Text), 11, 3)
    xxBlok = Mid(Trim(aNOP.Text), 15, 3)
    xxUrut = Mid(Trim(aNOP.Text), 19, 4)
    xxJenis = Right(Trim(aNOP.Text), 1)

    C_STR = "HAPUS_BANGUNAN '" & xxKec & "','" & xxKel & "','" & xxBlok & "','" & xxUrut & "','" & xxJenis & "'," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "'," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "'," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "'," & _
    "'" & xxJPB & "','" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "'," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "' ," & _
    "'" & Trim(aNOP.Text) & "','" & cboPajak(1).Text * 1 & "'"
    openDB (C_STR)
    MsgBox "Data Telah Berhasil Dihapus...", vbInformation, "Sukses...!"
    Log2
    cmdCancel_Click
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub


Private Sub cTPajak_Click()
On Error Resume Next
If cTPajak.Text * 1 < cboPajak(2).Text * 1 Then
    cboPajak(2).Text = cTPajak.Text
End If
'DBKB_FAS1
'DBKB_FAS3

End Sub

Private Sub cTPajak_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
        If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
         KeyAscii = 0
        End If
End Sub

Private Sub cTPajak_LostFocus()
On Error Resume Next
'If cTPajak.Text = "" Then cTPajak.Text = cTPajak.List(0)
For i = 0 To cTPajak.ListCount - 1
        If (UCase(cTPajak.List(i)) Like "*" + UCase(cTPajak.Text) + "*" = True) Then
            cTPajak.Text = cTPajak.List(i)
            GoTo Keluar
        End If
          If i = cTPajak.ListCount - 1 Then
            If UCase(cTPajak.List(i)) Like "*" + UCase(cTPajak.Text) + "*" = False Then
                cTPajak.Text = cTPajak.List(0)
                GoTo Keluar
            End If
        End If
    Next
Keluar:
If cTPajak.Text * 1 < cboPajak(2).Text * 1 Then
    cboPajak(2).Text = cTPajak.Text
End If

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
'Me.Width = 8670
'Me.Height = 7155
'Me.Width = 10455

Me.Width = 10395
chEdit.Value = 0
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2

QBNG = "SELECT THN_AWAL_KLS_BNG FROM KELAS_BANGUNAN order by THN_AWAL_KLS_BNG desc"
openDB (QBNG)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
xTB = rPajak!THN_AWAL_KLS_BNG
'MsgBox xID
If xID = "" Then
For i = 0 To 2
        dtPajak(i).Value = Format(Now, "dd/mm/yyyy")
    Next
cboPajak(0).Clear: cboPajak(1).Clear: cboPajak(2).Clear: cboPajak(3).Clear
cboPajak(4).Clear: cboPajak(5).Clear: cboPajak(6).Clear
cboPajak(7).Clear: cboPajak(8).Clear: cboPajak(9).Clear
cboPajak(1).Text = 1
cboPajak(2).Text = Format(Now, "yyyy")
cboPajak(3).Text = 0 'Format(Now, "yyyy")
cStandard.Value = 0
For i = 1 To 100
    cboPajak(1).AddItem i
Next
cboPajak(3).AddItem "0"
For i = 1 To 100
    cboPajak(2).AddItem Format(Now, "yyyy") + 1 - i
    cboPajak(3).AddItem Format(Now, "yyyy") + 1 - i
Next

cboPajak(4).Text = "02-Baik"
cboPajak(4).AddItem "01-Sangat Baik"
cboPajak(4).AddItem "02-Baik"
cboPajak(4).AddItem "03-Sedang"
cboPajak(4).AddItem "04-Jelek"

cboPajak(5).Text = "02-Beton"
cboPajak(5).AddItem "01-Baja"
cboPajak(5).AddItem "02-Beton"
cboPajak(5).AddItem "03-Batu Bata"
cboPajak(5).AddItem "04-Kayu"

'cboPajak(6).Text = "01-Decrabon/Beton/Gtg Glazur"
'cboPajak(6).AddItem "01-Decrabon/Beton/Gtg Glazur"
'cboPajak(6).AddItem "02-Gtg Beton/Aluminium"
'cboPajak(6).AddItem "03-Gtg Biasa/Sirap"
'cboPajak(6).AddItem "04-Asbes"
'cboPajak(6).AddItem "05-Seng"
Call Tampil_Material("23", 6)
cboPajak(6).Text = cboPajak(6).List(4)
'cboPajak(7).Text = "01-Kaca/Aluminium"
'cboPajak(7).AddItem "01-Kaca/Aluminium"
'cboPajak(7).AddItem "02-Beton"
'cboPajak(7).AddItem "03-Batu Bata/Conblok"
'cboPajak(7).AddItem "04-Kayu"
'cboPajak(7).AddItem "05-Seng"
'cboPajak(7).AddItem "06-Spandex"
Call Tampil_Material("21", 7)
cboPajak(7).Text = cboPajak(7).List(1)
xxNo = cboPajak(7).ListCount
cboPajak(7).AddItem xxNo + 1 & " TIDAK ADA"
'cboPajak(8).Text = "01-Marmer"
'cboPajak(8).AddItem "01-Marmer"
'cboPajak(8).AddItem "02-Keramik"
'cboPajak(8).AddItem "03-Teraso"
'cboPajak(8).AddItem "04-Ubin PC/Papan"
'cboPajak(8).AddItem "05-Semen"
Call Tampil_Material("22", 8)
cboPajak(8).Text = cboPajak(8).List(4)

'cboPajak(9).Text = "01-Akuistik/Jati"
'cboPajak(9).AddItem "01-Akuistik/Jati"
'cboPajak(9).AddItem "02-Triplek/Asbes/Bambu"
'cboPajak(9).AddItem "30-Tidak Ada"
Call Tampil_Material("24", 9)
cboPajak(9).Text = cboPajak(9).List(1)
yyNO = cboPajak(9).ListCount
cboPajak(9).AddItem yyNO + 1 & " TIDAK ADA"
tampil_JPB
If xxNon = 0 Then
    cboPajak(0).Text = cboPajak(0).List(0)
Else
    cboPajak(0).Text = xxJPB
End If

'cboPajak(0).AddItem "01-Perumahan"
'cboPajak(0).AddItem "02-Perkantoran"
'cboPajak(0).AddItem "03-Pabrik"
'cboPajak(0).AddItem "04-Toko/Apotik/Pasar/Ruko"
'cboPajak(0).AddItem "05-Rumah Sakit/Klinik"
'cboPajak(0).AddItem "06-Olah Raga/Rekreasi"
'cboPajak(0).AddItem "07-Hotel/Wisma"
'cboPajak(0).AddItem "08-Bengkel/Gudang/Pertanian"
'cboPajak(0).AddItem "09-Gedung Pemerintahan"
'cboPajak(0).AddItem "10-Lain-Lain"
'cboPajak(0).AddItem "11-Bangunan Tidak Kena Pajak"
'cboPajak(0).AddItem "12-Bangunan Parkir"
'cboPajak(0).AddItem "13-Apartemen"
'cboPajak(0).AddItem "14-Pompa Bensin"
'cboPajak(0).AddItem "15-Tangki Minyak"
'cboPajak(0).AddItem "16-Gedung Sekolah"


cboPajak(17).Clear
cboPajak(17).Text = "02-Tidak Ada"
cboPajak(17).AddItem "01-Ada"
cboPajak(17).AddItem "02-Tidak Ada"
cboPajak(16).Clear
cboPajak(16).Text = "01-Baja/Besi"
cboPajak(16).AddItem "01-Baja/Besi"
cboPajak(16).AddItem "02-Bata/Batako"
For i = 13 To 15
    cboPajak(i).Clear
Next
For i = 13 To 15
    cboPajak(i).Text = "02-Tidak Ada"
    cboPajak(i).AddItem "01-Ada"
    cboPajak(i).AddItem "02-Tidak Ada"
Next
cboPajak(19).Clear
cboPajak(19).Text = "01-Diplester"
cboPajak(19).AddItem "01-Diplester"
cboPajak(19).AddItem "02-Dengan Pelapis"
If chPajak(3).Value = 0 Then
    txtPajak(6).Enabled = False
    txtPajak(6).BackColor = vbButtonFace
Else
    txtPajak(6).Enabled = True
    txtPajak(6).BackColor = vbWhite
End If

For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = 0
    End If
    bersih
Next
tKet.Text = "-"
'If frmObjek_Pajak_Bm.txtPajak(1).Text = "" Or frmObjek_Pajak_Bm.txtPajak(1).Text = "0" Then txtPajak(0).Text = 0 Else txtPajak(0).Text = frmObjek_Pajak_Bm.txtPajak(1).Text
'MsgBox NO_FORM
'If NO_FORM <> "" Or NO_FORM <> 0 Then

For i = 7 To 9
    txtPajak(i).Text = "-"
Next


cTPajak.Clear
cTPajak.Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    cTPajak.AddItem i
Next
txtPajak(4).Text = 1
chPajak(1).Value = 1
End If
'txtPajak(1).Text = frmObjek_Pajak_Bm.tBumi(0).Text
'DBKB_FAS1
'DBKB_FAS3

'frmObjek_Pajak_Bm.Hide

If byPass = "01" Or byPass = "04" Then
    'txtPajak(0).Text = NO_FORM
    'txtPajak(0).Locked = True
    'txtPajak(0).Enabled = False
    If bypass4 = 2 Then
        aNOP.Enabled = True
        cmdNOP.Enabled = True
        cTPajak.Enabled = True
        chPajak(0).Value = 0
        chPajak(1).Value = 1
        chPajak(2).Value = 0
        chPajak(0).Enabled = False
        chPajak(1).Enabled = True
        chPajak(2).Enabled = True
        
    Else
        aNOP.Enabled = False
        cmdNOP.Enabled = False
        cTPajak.Enabled = False
        chPajak(0).Value = 1
        chPajak(1).Value = 0
        chPajak(2).Value = 0
        chPajak(1).Enabled = False
        chPajak(2).Enabled = False
        
    End If
    txtPajak(1).Text = BYPASS1
    cTPajak.Text = BYPASS2
    cboPajak(2).Text = BYPASS2
    txtPajak(2).Text = BYPASS3
   'LID.Caption = BYPASS3
    'MsgBox byPass
'Else 'If NO_FORM = 0 Then
'    txtPajak(0).Locked = False
'    txtPajak(0).Enabled = True
'    aNOP.Enabled = True
'    cmdNOP.Enabled = True
End If
'If chPajak(1).Value = 1 Or chPajak(2).Value = 1 Then tempLog1
Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
xxNon = 0
xID = ""
byPass = "03"
End Sub

Sub callDBKBUtama()
On Error GoTo Salah
Dim xNilai(10)
Dim xJum
JPB = Left(Trim(cboPajak(0).Text), 2)
If txtPajak(4).Text > 1 Then
    JL = 2
Else
    JL = 1
End If
'STRQ = "Select * From DBKB_STANDARD where THN_DBKB_STANDARD ='" & cTPajak.Text*1-1 & "' and KD_JPB='" & JPB & "' and KD_BNG_LANTAI = '" & JPB & "_" & txtPajak(4).Text & "_" & cTipe & "' order BY TIPE_BNG ASC"
StrQ = "Select * From DBKB_STANDARD where THN_DBKB_STANDARD ='" & cTPajak.Text * 1 & "' and KD_JPB='" & JPB & "'  order BY KD_BNG_LANTAI,kd_JPB,TIPE_BNG ASC"
openDB (StrQ)
MsgBox Format(cTipe, "000")
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        'tes
        xJum = 0
Do While Not rPajak.EOF
    'cKode = Trim(rPajak!KD_BNG_LANTAI)
    If Trim(rPajak!TIPE_BNG) = Format(cTipe, "000") Then
        'xKode = JPB * 1 & "_" & JL & "_" & Format(cTipe, "000")
        xJum = xJum + 1
        xNilai(xJum) = rPajak!NILAI_DBKB_STANDARD
    
    End If
'    xKode = JPB * 1 & "_1_" & Format(cTipe, "000")
'    xKode1 = JPB * 1 & "_2_" & Format(cTipe, "000")
'    If xKode = Trim(rPajak!KD_BNG_LANTAI) Then
'       LDBKB.Caption = rPajak!NILAI_DBKB_STANDARD
'       GoTo cetak
'    ElseIf xKode1 = Trim(rPajak!KD_BNG_LANTAI) Then
'        LDBKB.Caption = rPajak!NILAI_DBKB_STANDARD
'
'    End If
    rPajak.MoveNext
Loop
If xJum = 1 Then
    nDBKB = xNilai(xJum)
Else
    If JL = 1 Then
        nDBKB = xNilai(1)
    Else
        nDBKB = xNilai(2)
    End If
End If
If nDBKB = "" Then
    If JPB = 11 Then
        nDBKB = 0
    Else
        MsgBox "Ubah Jenis Penggunaan Bangunan.."
    End If
Else
    LDBKB.Caption = nDBKB
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_DINDING(xKerja, xNama)
On Error GoTo Salah
StrQ = "SELECT * FROM VMATERIAL WHERE THN_DBKB_MATERIAL ='" & cTPajak.Text * 1 & "' and (KD_PEKERJAAN='" & xKerja & "' and NM_KEGIATAN='" & xNama & "') ORDER BY KD_PEKERJAAN,KD_KEGIATAN ASC"
openDB (StrQ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
        xDinding = rPajak!NILAI_DBKB_MATERIAL
rPajak.MoveNext
Loop
If IsNull(xDinding) = True Or xDinding = "" Or cboPajak(7).Text = cboPajak(7).List(cboPajak(7).ListCount - 1) Then
    xDinding = 0
End If

LDinding.Caption = xDinding
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_LANTAI(xKerja, xNama)
On Error GoTo Salah
StrQ = "SELECT * FROM VMATERIAL WHERE THN_DBKB_MATERIAL ='" & cTPajak.Text * 1 & "' and (KD_PEKERJAAN='" & xKerja & "' and NM_KEGIATAN='" & xNama & "') ORDER BY KD_PEKERJAAN,KD_KEGIATAN ASC"
openDB (StrQ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
        xLantai = rPajak!NILAI_DBKB_MATERIAL
rPajak.MoveNext
Loop
LLantai.Caption = xLantai
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_ATAP(xKerja, xNama)
On Error GoTo Salah
StrQ = "SELECT * FROM VMATERIAL WHERE THN_DBKB_MATERIAL ='" & cTPajak.Text * 1 & "' and (KD_PEKERJAAN='" & xKerja & "' and NM_KEGIATAN='" & xNama & "') ORDER BY KD_PEKERJAAN,KD_KEGIATAN ASC"
openDB (StrQ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
        xAtap = rPajak!NILAI_DBKB_MATERIAL / txtPajak(4).Text
rPajak.MoveNext
Loop
LAtap.Caption = xAtap
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_LANGIT2(xKerja, xNama)
On Error GoTo Salah
StrQ = "SELECT * FROM VMATERIAL WHERE THN_DBKB_MATERIAL ='" & cTPajak.Text * 1 & "' and (KD_PEKERJAAN='" & xKerja & "' and NM_KEGIATAN='" & xNama & "') ORDER BY KD_PEKERJAAN,KD_KEGIATAN ASC"
'STRQ = "SELECT * FROM VMATERIAL WHERE THN_DBKB_MATERIAL ='" & cTPajak.Text * 1 & "' ORDER BY KD_PEKERJAAN,KD_KEGIATAN ASC"
openDB (StrQ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    
    xLangit2 = rPajak!NILAI_DBKB_MATERIAL
rPajak.MoveNext
Loop
If IsNull(xLangit2) = True Or xLangit2 = "" Or cboPajak(9).Text = cboPajak(9).List(cboPajak(9).ListCount - 1) Then
    xLangit2 = 0
End If
LLangit2.Caption = xLangit2
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub Tampil_Material(xKode As String, xxKode As Integer)
On Error GoTo Salah
StrQ = "SELECT * FROM VMATERIAL WHERE THN_DBKB_MATERIAL ='" & 2013 * 1 & "' and KD_PEKERJAAN= '" & xKode & " ' ORDER BY KD_PEKERJAAN,KD_KEGIATAN ASC"
openDB (StrQ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
        cboPajak(xxKode).AddItem rPajak!NM_KEGIATAN
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub tampil_JPB()
On Error GoTo Salah
StrQ = "SELECT * FROM REF_JPB order by KD_JPB"
openDB (StrQ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
        cboPajak(0).AddItem rPajak!KD_JPB & " " & rPajak!NM_JPB
rPajak.MoveNext
Loop

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub tampil_Susut(Nil_Bng)
On Error GoTo Salah
Dim xKondisi ', xKonstruksi, xDinding, xLangit2
xKondisi = Left(cboPajak(4).Text, 2) * 1
'xKonstruksi = Left(cboPajak(5).Text, 2) * 1
'xDinding = Left(cboPajak(7).Text, 1) * 1
'xLangit2 = Left(cboPajak(9).Text, 2) * 1
'MsgBox xKondisi & ":" & xKonstruksi & ":" & xDinding & ":" & xLangit2
StrQ = "SELECT PENYUSUTAN.KD_RANGE_PENYUSUTAN, PENYUSUTAN.UMUR_EFEKTIF, PENYUSUTAN.KONDISI_BNG_SUSUT, RANGE_PENYUSUTAN.NILAI_MIN_PENYUSUTAN, RANGE_PENYUSUTAN.NILAI_MAX_PENYUSUTAN, PENYUSUTAN.NILAI_PENYUSUTAN FROM PENYUSUTAN INNER JOIN RANGE_PENYUSUTAN ON PENYUSUTAN.KD_RANGE_PENYUSUTAN = RANGE_PENYUSUTAN.KD_RANGE_PENYUSUTAN ORDER BY PENYUSUTAN.UMUR_EFEKTIF" ' where PENYUSUTAN.KONDISI_BNG_SUSUT*1='" & xKondisi & "' and PENYUSUTAN.UMUR_EFEKTIF='" & umur_EFF & "' ORDER BY PENYUSUTAN.UMUR_EFEKTIF"
'STRQ = "SELECT PENYUSUTAN.KD_RANGE_PENYUSUTAN, PENYUSUTAN.UMUR_EFEKTIF, PENYUSUTAN.KONDISI_BNG_SUSUT, RANGE_PENYUSUTAN.NILAI_MIN_PENYUSUTAN, RANGE_PENYUSUTAN.NILAI_MAX_PENYUSUTAN, PENYUSUTAN.NILAI_PENYUSUTAN FROM PENYUSUTAN INNER JOIN RANGE_PENYUSUTAN ON PENYUSUTAN.KD_RANGE_PENYUSUTAN = RANGE_PENYUSUTAN.KD_RANGE_PENYUSUTAN ORDER BY PENYUSUTAN.UMUR_EFEKTIF ASC"
openDB (StrQ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'MsgBox rPajak![PENYUSUTAN.UMUR_EFEKTIF] & ":" & rPajak![PENYUSUTAN.KONDISI_BNG_SUSUT]
i = 0

Do While Not rPajak.EOF

If rPajak![UMUR_EFEKTIF] = umur_EFF And xKondisi = rPajak![KONDISI_BNG_SUSUT] * 1 Then
'    If xKonstruksi = 4 And (xDinding = 4 Or xDinding = 5) Then
'        If xDinding = 4 And xLangit2 = 3 Then
'            If nDBKB * 1000 >= rPajak![NILAI_MIN_PENYUSUTAN] And nDBKB * 1000 <= rPajak![NILAI_MAX_PENYUSUTAN] Then
'                xSUSUT = rPajak![NILAI_PENYUSUTAN]
'            End If
'        Else
'            If nDBKB * 1000000 >= rPajak![NILAI_MIN_PENYUSUTAN] And nDBKB * 1000000 <= rPajak![NILAI_MAX_PENYUSUTAN] Then
'                xSUSUT = rPajak![NILAI_PENYUSUTAN]
'            End If
'        End If
'    Else
        If Nil_Bng >= rPajak![NILAI_MIN_PENYUSUTAN] And Nil_Bng <= rPajak![NILAI_MAX_PENYUSUTAN] Then
                xSUSUT = rPajak![NILAI_PENYUSUTAN]
        End If
'    End If
End If
rPajak.MoveNext
Loop
LPersen.Caption = xSUSUT
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub DBKB_FAS()
On Error GoTo Salah
Dim NFAS
StrQ = "SELECT * FROM VKOMPONEN WHERE THN_HRG_RESOURCE ='" & cTPajak.Text * 1 & "' and KD_GROUP_RESOURCE >='33' ORDER BY KD_GROUP_RESOURCE,KD_RESOURCE ASC"
openDB (StrQ)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0: J = 0
Do While Not rPajak.EOF
    'If Trim(rPajak!KD_GROUP_RESOURCE) = "41" And Trim(rPajak!KD_RESOURCE) = "01" Then
    '    DAYA_LISTRIK = rPajak!HRG_RESOURECE
    'End If
    NFAS = rPajak!HRG_RESOURCE
    
    'Daya Listrik
    If Trim(rPajak!KD_GROUP_RESOURCE) = "41" And Trim(rPajak!KD_RESOURCE) = "01" Then
        'DAYA_LISTRIK = nFAS
    End If
    'AC SPLIT dan Window
    If Trim(rPajak!KD_GROUP_RESOURCE) = "37" Then
        If Trim(rPajak!KD_RESOURCE) = "16" Then
            JUM_SPLIT = NFAS
        Else
            JUM_WINDOW = NFAS
        End If
    End If
    'Pekerasan Halaman
    If Trim(rPajak!KD_GROUP_RESOURCE) = "47" Then
        If Trim(rPajak!KD_RESOURCE) = "01" Then
            LUAS_HRINGAN = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "02" Then
            LUAS_HSEDANG = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "03" Then
            LUAS_HBERAT = NFAS
        Else
            LUAS_HPENUTUP = NFAS
        End If
    End If
    
    'Lapangan Tenis
    If Trim(rPajak!KD_GROUP_RESOURCE) = "35" Then
      
        If Trim(rPajak!KD_RESOURCE) = "01" Then
            JUM_LAP_BETON1 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "02" Then
            JUM_LAP_ASPAL1 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "03" Then
            JUM_LAP_RUMPUT1 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "04" Then
            JUM_LAP_BETON2 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "05" Then
            JUM_LAP_ASPAL2 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "06" Then
            JUM_LAP_RUMPUT2 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "07" Then
            JUM_LAP_BETON11 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "08" Then
            JUM_LAP_ASPAL11 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "09" Then
            JUM_LAP_RUMPUT11 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "10" Then
            JUM_LAP_BETON21 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "11" Then
            JUM_LAP_ASPAL21 = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "12" Then
            JUM_LAP_RUMPUT21 = NFAS
        End If
    End If
    
    'Pagar
    If Trim(rPajak!KD_GROUP_RESOURCE) = "43" Then
        If Trim(rPajak!KD_RESOURCE) = "01" Then
            BAHAN_PAGAR1 = NFAS
        Else
            BAHAN_PAGAR2 = NFAS
        End If
    End If
    'Tangga Berjalan
    
    If Trim(rPajak!KD_GROUP_RESOURCE) = "39" Then
        If Trim(rPajak!KD_RESOURCE) = "01" Then
            LEBAR_TANGGA1 = NFAS
        Else
             LEBAR_TANGGA2 = NFAS
        End If
    End If
    
    'Pemadam Kebakaran
    If Trim(rPajak!KD_GROUP_RESOURCE) = "45" Then
        If Trim(rPajak!KD_RESOURCE) = "01" Then
            BAKAR_H = NFAS
        ElseIf Trim(rPajak!KD_RESOURCE) = "02" Then
            BAKAR_S = NFAS
        Else
            BAKAR_F = NFAS
        End If
    End If
    'PABX
    If Trim(rPajak!KD_GROUP_RESOURCE) = "44" Then
        If Trim(rPajak!KD_RESOURCE) = "01" Then
            JUM_PABX = NFAS
        Else
            JUM_PABX = 0
        End If
    End If
    'sumur
    If Trim(rPajak!KD_GROUP_RESOURCE) = "46" Then
        If Trim(rPajak!KD_RESOURCE) = "01" Then
            DALAM_SUMUR = NFAS
        Else
            DALAM_SUMUR = 0
        End If
    End If
    'Kolam Renang
    If Trim(rPajak!KD_GROUP_RESOURCE) = "33" Then
        i = i + 1
        Luas_Kolam(i) = NFAS
        
    End If
    
    'LIFT
    If Trim(rPajak!KD_GROUP_RESOURCE) = "38" Then
        J = J + 1
        JLIFT(J) = NFAS
    End If
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Sub callDBKBUtama1()
On Error GoTo Salah
Dim xNilai(10)
Dim xJum
JPB = Left(Trim(cboPajak(0).Text), 2)
If txtPajak(4).Text > 1 Then
    JL = 2
Else
    JL = 1
End If
'STRQ = "Select * From DBKB_STANDARD where THN_DBKB_STANDARD ='" & cTPajak.Text*1-1 & "' and KD_JPB='" & JPB & "' and KD_BNG_LANTAI = '" & JPB & "_" & txtPajak(4).Text & "_" & cTipe & "' order BY TIPE_BNG ASC"
StrQ = "SELECT DBKB_STANDARD.THN_DBKB_STANDARD, DBKB_STANDARD.KD_JPB, TIPE_BANGUNAN.TIPE_BNG, TIPE_BANGUNAN.NM_TIPE_BNG, TIPE_BANGUNAN.LUAS_MIN_TIPE_BNG, TIPE_BANGUNAN.LUAS_MAX_TIPE_BNG, TIPE_BANGUNAN.FAKTOR_PEMBAGI_TIPE_BNG, DBKB_STANDARD.KD_BNG_LANTAI, DBKB_STANDARD.NILAI_DBKB_STANDARD FROM DBKB_STANDARD INNER JOIN TIPE_BANGUNAN ON DBKB_STANDARD.TIPE_BNG = TIPE_BANGUNAN.TIPE_BNG WHERE (((DBKB_STANDARD.THN_DBKB_STANDARD)='" & cTPajak.Text * 1 & "'))"
openDB (StrQ)
MsgBox Format(cTipe, "000")
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        'tes
        xJum = 0
Do While Not rPajak.EOF
    'cKode = Trim(rPajak!KD_BNG_LANTAI)
    If Trim(rPajak!TIPE_BNG) = Format(cTipe, "000") Then
        'xKode = JPB * 1 & "_" & JL & "_" & Format(cTipe, "000")
        xJum = xJum + 1
        xNilai(xJum) = rPajak!NILAI_DBKB_STANDARD
    
    End If
'    xKode = JPB * 1 & "_1_" & Format(cTipe, "000")
'    xKode1 = JPB * 1 & "_2_" & Format(cTipe, "000")
'    If xKode = Trim(rPajak!KD_BNG_LANTAI) Then
'       LDBKB.Caption = rPajak!NILAI_DBKB_STANDARD
'       GoTo cetak
'    ElseIf xKode1 = Trim(rPajak!KD_BNG_LANTAI) Then
'        LDBKB.Caption = rPajak!NILAI_DBKB_STANDARD
'
'    End If
    rPajak.MoveNext
Loop
If xJum = 1 Then
    nDBKB = xNilai(xJum)
Else
    If JL = 1 Then
        nDBKB = xNilai(1)
    Else
        nDBKB = xNilai(2)
    End If
End If
If nDBKB = "" Then
    If JPB = 11 Then
        nDBKB = 0
    Else
        MsgBox "Ubah Jenis Penggunaan Bangunan.."
    End If
Else
    LDBKB.Caption = nDBKB
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub DBKB_FAS1()
On Error GoTo Salah
Dim NFAS

QSTR = "SELECT FASILITAS.KD_FASILITAS, FASILITAS.NM_FASILITAS, FASILITAS.SATUAN_FASILITAS, FAS_NON_DEP.NILAI_NON_DEP, FAS_NON_DEP.THN_NON_DEP FROM FASILITAS INNER JOIN FAS_NON_DEP ON FASILITAS.KD_FASILITAS = FAS_NON_DEP.KD_FASILITAS WHERE FAS_NON_DEP.THN_NON_DEP='" & cTPajak.Text & "' "

openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0: J = 0
Do While Not rPajak.EOF
    'If Trim(rPajak!KD_GROUP_RESOURCE) = "41" And Trim(rPajak!KD_RESOURCE) = "01" Then
    '    DAYA_LISTRIK = rPajak!HRG_RESOURECE
    'End If
    NFAS = rPajak!NILAI_NON_DEP
    
    'Daya Listrik
    If Trim(rPajak!KD_FASILITAS) = "44" Or UCase(Trim(rPajak!NM_FASILITAS)) = "LISTRIK" Then
        DAYA_LISTRIK = NFAS
    End If
    S_Listrik = Format(DAYA_LISTRIK, "#,#0.00")
    'AC SPLIT dan Window
    If Trim(rPajak!KD_FASILITAS) = "01" Then JUM_SPLIT = NFAS
    If Trim(rPajak!KD_FASILITAS) = "02" Then JUM_WINDOW = NFAS
    'AC Central Bangunan Lain
    
    If Trim(rPajak!KD_FASILITAS) = "11" Then JUM_AC_CENTRAL = NFAS
    
    'Pekerasan Halaman
    If Trim(rPajak!KD_FASILITAS) = "14" Then LUAS_HRINGAN = NFAS
    If Trim(rPajak!KD_FASILITAS) = "15" Then LUAS_HSEDANG = NFAS
    If Trim(rPajak!KD_FASILITAS) = "16" Then LUAS_HBERAT = NFAS
    If Trim(rPajak!KD_FASILITAS) = "17" Then LUAS_HPENUTUP = NFAS
    
    
    
    'Lapangan Tenis
      
        If Trim(rPajak!KD_FASILITAS) = "18" Then JUM_LAP_BETON1 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "19" Then JUM_LAP_ASPAL1 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "20" Then JUM_LAP_RUMPUT1 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "21" Then JUM_LAP_BETON2 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "22" Then JUM_LAP_ASPAL2 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "23" Then JUM_LAP_RUMPUT2 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "24" Then JUM_LAP_BETON11 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "25" Then JUM_LAP_ASPAL11 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "26" Then JUM_LAP_RUMPUT11 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "27" Then JUM_LAP_BETON21 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "28" Then JUM_LAP_ASPAL21 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "29" Then JUM_LAP_RUMPUT21 = NFAS
    
    'Pagar
    If Trim(rPajak!KD_FASILITAS) = "35" Then BAHAN_PAGAR1 = NFAS
    If Trim(rPajak!KD_FASILITAS) = "36" Then BAHAN_PAGAR2 = NFAS
    
    'Tangga Berjalan
    
        If Trim(rPajak!KD_FASILITAS) = "33" Then LEBAR_TANGGA1 = NFAS
        If Trim(rPajak!KD_FASILITAS) = "34" Then LEBAR_TANGGA2 = NFAS
    
    'Pemadam Kebakaran
        If Trim(rPajak!KD_FASILITAS) = "37" Then BAKAR_H = NFAS
        If Trim(rPajak!KD_FASILITAS) = "38" Then BAKAR_S = NFAS
        If Trim(rPajak!KD_FASILITAS) = "39" Then BAKAR_F = NFAS
    'PABX
        If Trim(rPajak!KD_FASILITAS) = "41" Then JUM_PABX = NFAS
        
    'sumur
        If Trim(rPajak!KD_FASILITAS) = "42" Then DALAM_SUMUR = NFAS
    'Kolam Renang
rPajak.MoveNext
Loop
'MsgBox BAKAR_H & ":" & BAKAR_S & ":" & BAKAR_F
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub DBKB_FAS3()
On Error GoTo Salah
Dim NFAS, xMIN, xMAX

QSTR = "SELECT FASILITAS.KD_FASILITAS, FASILITAS.NM_FASILITAS, FASILITAS.SATUAN_FASILITAS, FAS_DEP_MIN_MAX.KLS_DEP_MIN, FAS_DEP_MIN_MAX.KLS_DEP_MAX, FAS_DEP_MIN_MAX.NILAI_DEP_MIN_MAX, FAS_DEP_MIN_MAX.THN_DEP_MIN_MAX FROM FASILITAS INNER JOIN FAS_DEP_MIN_MAX ON FASILITAS.KD_FASILITAS = FAS_DEP_MIN_MAX.KD_FASILITAS WHERE FAS_DEP_MIN_MAX.THN_DEP_MIN_MAX='" & cTPajak.Text & "' ORDER BY FASILITAS.KD_FASILITAS,FAS_DEP_MIN_MAX.KLS_DEP_MIN,FAS_DEP_MIN_MAX.KLS_DEP_MAX ASC"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0: J = 0: A = 0: K = 0: L = 0: M = 0
Do While Not rPajak.EOF
    NFAS = rPajak!NILAI_DEP_MIN_MAX
    xMIN = rPajak!KLS_DEP_MIN
    xMAX = rPajak!KLS_DEP_MAX
    'Kolam Renang
    If Trim(rPajak!KD_FASILITAS) = "12" And Left(cboPajak(19).Text, 2) = "01" Then
            If tKolam.Text * 1 >= xMIN And tKolam.Text * 1 <= xMAX Then
                Luas_Kolam = NFAS
            End If
    End If
    If Trim(rPajak!KD_FASILITAS) = "13" And Left(cboPajak(19).Text, 2) = "02" Then
            If tKolam.Text * 1 >= xMIN And tKolam.Text * 1 <= xMAX Then
                Luas_Kolam = NFAS
            End If
    End If
     
    'LIFT
    
    If Trim(rPajak!KD_FASILITAS) = "30" Then
        If tLift1.Text * 1 >= xMIN And tLift1.Text * 1 <= xMAX Then
            JLIFT(1) = NFAS
        End If
   ElseIf Trim(rPajak!KD_FASILITAS) = "31" Then
        If tLift2.Text * 1 >= xMIN And tLift2.Text * 1 <= xMAX Then
            JLIFT(2) = NFAS
        End If
    End If
    If Trim(rPajak!KD_FASILITAS) = "32" Then
        If tLift3.Text * 1 >= xMIN And tLift3.Text * 1 <= xMAX Then
            JLIFT(3) = NFAS
        End If
    End If
    'Genset
    
     If Trim(rPajak!KD_FASILITAS) = "40" Then
        If tGenset.Text * 1 >= xMIN And tGenset.Text * 1 <= xMAX Then
            JUM_GENSET = NFAS
        End If
    End If
rPajak.MoveNext
Loop

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub DBKB_FAS2()
On Error GoTo Salah
Dim NFAS, xKelas
QSTR = "SELECT FASILITAS.KD_FASILITAS, FASILITAS.NM_FASILITAS, FASILITAS.SATUAN_FASILITAS, FAS_DEP_JPB_KLS_BINTANG.KLS_BINTANG, FAS_DEP_JPB_KLS_BINTANG.NILAI_FASILITAS_KLS_BINTANG, FAS_DEP_JPB_KLS_BINTANG.THN_DEP_JPB_KLS_BINTANG FROM FASILITAS INNER JOIN FAS_DEP_JPB_KLS_BINTANG ON FASILITAS.KD_FASILITAS = FAS_DEP_JPB_KLS_BINTANG.KD_FASILITAS WHERE FAS_DEP_JPB_KLS_BINTANG.THN_DEP_JPB_KLS_BINTANG='" & cTPajak.Text & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    NFAS = rPajak!NILAI_FASILITAS_KLS_BINTANG
    xKelas = Trim(rPajak!KLS_BINTANG)
    'Boiler Hotel
    If Trim(rPajak!KD_FASILITAS) = "43" Then
            If cJPB7b.Text = xKelas Then
                Nil_Boiler_Ht = NFAS
            End If
    End If
    'Boiler Apartemen
    If Trim(rPajak!KD_FASILITAS) = "45" Then
            If cJPB13.Text = xKelas Then
                Nil_Boiler_Ap = NFAS
            End If
    End If
    'AC Central Kantor JPB=2
    If Trim(rPajak!KD_FASILITAS) = "03" Then
            If cJPB29.Text = xKelas Then
                Nil_AC_Central(1) = NFAS
            End If
    End If
    'AC Central Kamar Hotel JPB = 7
        If cJPB7a.Text = xKelas Then
            If Trim(rPajak!KD_FASILITAS) = "04" Then
                Nil_AC_Central(2) = NFAS * JPB7b.Text 'Kamar Hotel
            ElseIf Trim(rPajak!KD_FASILITAS) = "05" Then
                Nil_AC_Central(3) = NFAS * JPB7c.Text 'Ruangan Lain
            End If
        End If
    'AC Central Pertokoan JPB = 4
        If Trim(rPajak!KD_FASILITAS) = "06" Then
            If cJPB4.Text = xKelas Then
                Nil_AC_Central(4) = NFAS
            End If
        End If
    'AC Central Kamar Rumah Sakit JPB = 5
        
            If cJPB5.Text = xKelas Then
                If Trim(rPajak!KD_FASILITAS) = "07" Then
                    Nil_AC_Central(5) = NFAS * JPB5a.Text * 1 'Ruangan Rumah Sakit
                ElseIf Trim(rPajak!KD_FASILITAS) = "08" Then
                    Nil_AC_Central(6) = NFAS * JPB5b.Text 'Ruangan Lain
                End If
            End If
    'AC Central Apartemen JPB = 13
        
            If cJPB13.Text = xKelas Then
                If Trim(rPajak!KD_FASILITAS) = "09" Then
                    Nil_AC_Central(7) = NFAS * JPB13c.Text * 1 'Kamar Apartemen
                ElseIf Trim(rPajak!KD_FASILITAS) = "10" Then
                    Nil_AC_Central(8) = NFAS * JPB13a.Text * 1 'Ruangan Lain Apartemen
                End If
            End If
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub


Sub call_Mezz()
On Error GoTo Salah
QSTR = "SELECT * FROM DBKB_MEZANIN WHERE THN_DBKB_MEZANIN='" & cTPajak.Text & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    If JPB38(4).Text = "" Then JPB38(4).Text = 0
    nMezanin = rPajak!NILAI_DBKB_MEZANIN '* JPB38(4).Text
    rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub call_Dukung()
On Error GoTo Salah
QSTR = "SELECT DBKB_DAYA_DUKUNG.KD_PROPINSI, DBKB_DAYA_DUKUNG.KD_DATI2, DBKB_DAYA_DUKUNG.THN_DBKB_DAYA_DUKUNG, DBKB_DAYA_DUKUNG.TYPE_KONSTRUKSI, DAYA_DUKUNG.DAYA_DUKUNG_LANTAI_MIN_DBKB, DAYA_DUKUNG.DAYA_DUKUNG_LANTAI_MAX_DBKB, DBKB_DAYA_DUKUNG.NILAI_DBKB_DAYA_DUKUNG FROM DBKB_DAYA_DUKUNG INNER JOIN DAYA_DUKUNG ON DBKB_DAYA_DUKUNG.TYPE_KONSTRUKSI = DAYA_DUKUNG.TYPE_KONSTRUKSI WHERE DBKB_DAYA_DUKUNG.THN_DBKB_DAYA_DUKUNG='" & cTPajak.Text & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0
Do While Not rPajak.EOF
    If JPB38(2).Text * txtPajak(3).Text >= rPajak!DAYA_DUKUNG_LANTAI_MIN_DBKB And JPB38(2).Text * txtPajak(3).Text <= rPajak!DAYA_DUKUNG_LANTAI_MAX_DBKB Then
        i = i + 1
        nDUKUNG = rPajak!NILAI_DBKB_DAYA_DUKUNG * txtPajak(3).Text
        nTipe_K = rPajak!TYPE_KONSTRUKSI
    End If
    rPajak.MoveNext
Loop
If nTipe_K = "" Or IsNull(nTipe_K) Then nTipe_K = 1
'If I = 0 Then nDUKUNG = 0
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub J_Karakter()
On Error GoTo Salah
Dim jmlText, jmlChar, i As Integer
    jmlChar = 0
    jmlText = Len(txtPajak(1).Text)
    For i = 0 To jmlText
        txtPajak(1).SelStart = i
        txtPajak(1).SelLength = 1
        If txtPajak(1).SelText = "_" Then
            jmlChar = jmlChar + 1
        End If
    Next
    totChar = jmlChar
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub tAC1_GotFocus()
Call c_blok(tAC1)
End Sub

Private Sub tAC1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tAC1_LostFocus()
Call c_Kosong(tAC1)
End Sub

Private Sub tAC2_GotFocus()
Call c_blok(tAC2)
End Sub

Private Sub tAC2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tAC2_LostFocus()
Call c_Kosong(tAC2)
End Sub

Private Sub tGenset_GotFocus()
Call c_blok(tGenset)
End Sub

Private Sub tGenset_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tGenset_LostFocus()
Call c_Kosong(tGenset)
End Sub

Private Sub tHal1_GotFocus()
Call c_blok(tHal1)
End Sub

Private Sub tHal1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tHal1_LostFocus()
Call c_Kosong(tHal1)
End Sub

Private Sub tHal2_GotFocus()
Call c_blok(tHal2)
End Sub

Private Sub tHal2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tHal2_LostFocus()
Call c_Kosong(tHal2)
End Sub

Private Sub tHal3_GotFocus()
Call c_blok(tHal3)
End Sub

Private Sub tHal3_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tHal3_LostFocus()
Call c_Kosong(tHal3)
End Sub

Private Sub tHal4_GotFocus()
Call c_blok(tHal4)
End Sub

Private Sub tHal4_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tHal4_LostFocus()
Call c_Kosong(tHal4)
End Sub

Private Sub tKet_LostFocus()
tKet.Text = Rep(tKet.Text)
End Sub

Private Sub tKolam_GotFocus()
Call c_blok(tKolam)
End Sub

Private Sub tKolam_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tKolam_LostFocus()
Call c_Kosong(tKolam)
End Sub

Private Sub tLap_Aspal1_GotFocus()
Call c_blok(tLap_Aspal1)
End Sub

Private Sub tLap_Aspal1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tLap_Aspal1_LostFocus()
Call c_Kosong(tLap_Aspal1)
End Sub

Private Sub tLap_Aspal2_GotFocus()
Call c_blok(tLap_Aspal2)
End Sub

Private Sub tLap_Aspal2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tLap_Aspal2_LostFocus()
Call c_Kosong(tLap_Aspal2)
End Sub

Private Sub tLap_Beton1_GotFocus()
Call c_blok(tLap_Beton1)
End Sub

Private Sub tLap_Beton1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tLap_Beton1_LostFocus()
Call c_Kosong(tLap_Beton1)
End Sub

Private Sub tLap_Beton2_GotFocus()
Call c_blok(tLap_Beton2)
End Sub

Private Sub tLap_Beton2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tLap_Beton2_LostFocus()
Call c_Kosong(tLap_Beton2)
End Sub

Private Sub tLap_Tanah1_GotFocus()
Call c_blok(tLap_Tanah1)
End Sub

Private Sub tLap_Tanah1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tLap_Tanah1_LostFocus()
Call c_Kosong(tLap_Tanah1)
End Sub

Private Sub tLap_Tanah2_GotFocus()
Call c_blok(tLap_Tanah2)
End Sub

Private Sub tLap_Tanah2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tLap_Tanah2_LostFocus()
Call c_Kosong(tLap_Tanah2)
End Sub

Private Sub tLift1_GotFocus()
Call c_blok(tLift1)
End Sub

Private Sub tLift1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tLift1_LostFocus()
Call c_Kosong(tLift1)
End Sub

Private Sub tLift2_GotFocus()
Call c_blok(tLift2)
End Sub

Private Sub tLift2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tLift2_LostFocus()
Call c_Kosong(tLift2)
End Sub

Private Sub tLift3_GotFocus()
Call c_blok(tLift3)
End Sub

Private Sub tLift3_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tLift3_LostFocus()
Call c_Kosong(tLift3)
End Sub

Private Sub tListrik_GotFocus()
'SendKeys "{Home}+{end}"
'tListrik.SetFocus
'tListrik.SelStart = 0
'tListrik.SelLength = Len(tListrik.Text)
'tListrik.SetFocus
Call c_blok(tListrik)
End Sub

Private Sub tListrik_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tListrik_LostFocus()
'If tListrik.Text = "" Or tListrik.Text = "-" Or tListrik.Text = "." Then
'    tListrik.Text = 0
'End If
Call c_Kosong(tListrik)
End Sub

Private Sub tPABX_GotFocus()
Call c_blok(tPABX)
End Sub

Private Sub tPABX_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tPABX_LostFocus()
Call c_Kosong(tPABX)
End Sub

Private Sub tPagar_GotFocus()
Call c_blok(tPagar)
End Sub

Private Sub tPagar_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tPagar_LostFocus()
Call c_Kosong(tPagar)
End Sub

Private Sub tSumur_GotFocus()
Call c_blok(tSumur)
End Sub

Private Sub tSumur_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tSumur_LostFocus()
Call c_Kosong(tSumur)
End Sub

Private Sub tTangga1_GotFocus()
Call c_blok(tTangga1)
End Sub

Private Sub tTangga1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tTangga1_LostFocus()
Call c_Kosong(tTangga1)
End Sub

Private Sub tTangga2_GotFocus()
Call c_blok(tTangga2)
End Sub

Private Sub tTangga2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub tTangga2_LostFocus()
Call c_Kosong(tTangga2)
End Sub

Private Sub txtPajak_GotFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
    txtPajak(0).SelStart = 0
    txtPajak(0).SelLength = Len(txtPajak(0).Text)
    txtPajak(0).Alignment = 0

Case 3 To 9
    Call c_blok(txtPajak(Index))
End Select
End Sub

Private Sub txtPajak_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
Select Case Index
Case 0, 3 To 6
   ' If KeyAscii = 13 Then
    '    SendKeys "{tab}"
    'End If
    
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
    End If
End Select

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

Private Sub txtPajak_LostFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 5
    cmdHitung_Click
Case 0, 3 To 9
    txtPajak(Index).Text = Rep(txtPajak(Index).Text)
    Call c_Kosong(txtPajak(Index))
End Select
If txtPajak(4).Text * 1 > 4 Then cStandard.Value = 1 Else cStandard.Value = 0
End Sub
Sub bersih()
On Error Resume Next
cJPB29.Clear
cJPB29.Text = "1"
cJPB4.Clear
cJPB4.Text = "1"
cJPB5.Clear
cJPB5.Text = "1"
cJPB6.Clear
cJPB6.Text = "1"
cJPB7a.Clear
cJPB7a.Text = "1"
cJPB7b.Clear
cJPB7b.Text = "0"
cJPB12.Clear
cJPB12.Text = "1"
cJPB13.Clear
cJPB13.Text = "1"
cJPB15.Clear
cJPB15.Text = "1"
cJPB16.Clear
cJPB16.Text = "1"
For i = 0 To 4
    JPB38(i).Text = 0
Next
cJPB17.Text = 0: cJPB17b.Text = 0: JPB29.Text = 0: JPB4.Text = 0
JPB5a.Text = 0: JPB5b.Text = 0
JPB7a.Text = 0: JPB7b.Text = 0: JPB7c.Text = 0: JPB7d.Text = 0
JPB13a.Text = 0: JPB13b.Text = 0: JPB13c.Text = 0: JPB13d.Text = 0
JPB15.Text = 0
End Sub


Private Sub xUP_DownClick()
On Error GoTo Salah
If cboPajak(1).Text > 1 Then
    cboPajak(1).Text = cboPajak(1).Text - 1
End If
StrQ1 = "Select * From DAT_OP_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' ORDER BY NO_BNG*1 DESC"
openDB (StrQ1)

call_data
If chPajak(1).Value = 1 Or chPajak(2).Value = 1 Or xxLanjut <> 1 Then tempLog1
Me.Width = 10395
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub xUP_UpClick()
On Error GoTo Salah
    cboPajak(1).Text = cboPajak(1).Text + 1
    StrQ1 = "Select * From DAT_OP_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' ORDER BY NO_BNG*1 DESC"
    openDB (StrQ1)
    call_data
    If chPajak(1).Value = 1 Or chPajak(2).Value = 1 Then tempLog1
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
    
End Sub
Sub call_data()
On Error GoTo Salah
cmdCancel_Click
txtPajak(1).Text = aNOP.Text

xJPB = Left(Trim(cboPajak(0).Text), 2)

    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    
    If rPajak.EOF Then
        If chPajak(1).Value = 1 Or chPajak(2).Value = 1 Then
            J_Karakter
            If Len(Trim(txtPajak(1).Text)) - (totChar * 1) = 24 Then
                MsgBox "Data Tidak Dapat Ditemukan...", vbCritical, "Tetnong"
                cboPajak(1).Text = cboPajak(1).Text - 1
                If cboPajak(1).Text = 0 Then cboPajak(1).Text = 1
                StrQ1 = "Select * From DAT_OP_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "'  ORDER BY NO_BNG*1 DESC"
                openDB (StrQ1)
            End If
        End If
    Else
        If chPajak(0).Value = 1 Then
            MsgBox "Data Untuk Nomor Bangunan Yang Dipilih Sudah Ada, " & _
                    vbCrLf & "Silahkan Diganti...", vbCritical, "Tetnong..."
                    cboPajak(1).SetFocus
                    cboPajak(1).Text = cboPajak(1).Text + 1
            Exit Sub
        End If
    End If
    
    Do While Not rPajak.EOF
        txtPajak(0).Text = rPajak!NO_FORMULIR_LSPOP
        cboPajak(1).Text = rPajak!NO_BNG * 1
        txtPajak(4).Text = rPajak!JML_LANTAI_BNG
        txtPajak(3).Text = rPajak!LUAS_BNG
        SELISIH_LUAS_EDIT = rPajak!LUAS_BNG
        cboPajak(0).Text = cboPajak(0).List(rPajak!KD_JPB * 1 - 1)
        cboPajak(2).Text = rPajak!THN_DIBANGUN_BNG
        If IsNull(rPajak!THN_RENOVASI_BNG) = True Then
            cboPajak(3).Text = 0
        Else
            cboPajak(3).Text = rPajak!THN_RENOVASI_BNG
        End If
        cboPajak(4).Text = cboPajak(4).List(rPajak!KONDISI_BNG * 1 - 1)
        cboPajak(5).Text = cboPajak(5).List(rPajak!JNS_KONSTRUKSI_BNG * 1 - 1)
        cboPajak(6).Text = cboPajak(6).List(rPajak!JNS_ATAP_BNG * 1 - 1)
        cboPajak(7).Text = cboPajak(7).List(rPajak!KD_DINDING * 1 - 1)
        cboPajak(8).Text = cboPajak(8).List(rPajak!KD_LANTAI * 1 - 1)
        cboPajak(9).Text = cboPajak(9).List(rPajak!KD_LANGIT_LANGIT * 1 - 1)
        If IsNull(rPajak!NIP_PENDATA_BNG) = True Or rPajak!NIP_PENDATA_BNG = "" Then rPajak!NIP_PENDATA_BNG = "-"
        txtPajak(7) = rPajak!NIP_PENDATA_BNG 'NIP Pendata
        If IsNull(rPajak!NIP_PEMERIKSA_BNG) = True Or rPajak!NIP_PEMERIKSA_BNG = "" Then rPajak!NIP_PEMERIKSA_BNG = "-"
        txtPajak(8) = rPajak!NIP_PEMERIKSA_BNG 'NIP Pemeriksa
        If IsNull(rPajak!NIP_PEREKAM_BNG) = True Or rPajak!NIP_PEREKAM_BNG = "" Then rPajak!NIP_PEREKAM_BNG = "-"
        txtPajak(9) = rPajak!NIP_PEREKAM_BNG 'NIP Perekam
        
        dtPajak(0).Value = Format(rPajak!TGL_PENDATAAN_BNG, "DD/MM/YYYY")
        'txtPajak(7).Text = rPajak!NIP_PENDATA_BNG
        dtPajak(1).Value = Format(rPajak!TGL_PEMERIKSAAN_BNG, "DD/MM/YYYY")
        'txtPajak(8).Text = rPajak!NIP_PEMERIKSA_BNG
        dtPajak(2).Value = Format(rPajak!TGL_PEREKAMAN_BNG, "DD/MM/YYYY")
        'txtPajak(9).Text = rPajak!NIP_PEREKAM_BNG
        rPajak.MoveNext
    Loop
    'Panggil Data Fasilitas
    StrQ2 = "Select * From DAT_FASILITAS_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' ORDER BY NO_BNG*1 DESC"
    openDB (StrQ2)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        If rPajak!KD_FASILITAS = "35" Then
            cboPajak(16).Text = cboPajak(16).List(0)
            tPagar.Text = rPajak!JML_SATUAN
        ElseIf rPajak!KD_FASILITAS = "36" Then
            cboPajak(16).Text = cboPajak(16).List(1)
            tPagar.Text = rPajak!JML_SATUAN
        End If
        If rPajak!KD_FASILITAS = "14" Then tHal1.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "15" Then tHal2.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "16" Then tHal3.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "17" Then tHal4.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "40" Then tGenset.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "41" Then tPABX.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "42" Then tSumur.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "12" Then
            cboPajak(19).Text = cboPajak(19).List(0)
            tKolam.Text = rPajak!JML_SATUAN
        ElseIf rPajak!KD_FASILITAS = "13" Then
            cboPajak(19).Text = cboPajak(19).List(1)
            tKolam.Text = rPajak!JML_SATUAN
        End If
        If rPajak!KD_FASILITAS = "37" Then cboPajak(13).Text = cboPajak(13).List(0)
        If rPajak!KD_FASILITAS = "38" Then cboPajak(13).Text = cboPajak(15).List(0)
        If rPajak!KD_FASILITAS = "39" Then cboPajak(13).Text = cboPajak(14).List(0)
        
        If rPajak!KD_FASILITAS = "18" Then tLap_Beton1.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "19" Then tLap_Aspal1.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "20" Then tLap_Tanah1.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "21" Then tLap_Beton2.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "22" Then tLap_Aspal2.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "23" Then tLap_Tanah2.Text = rPajak!JML_SATUAN
        
        If rPajak!KD_FASILITAS = "24" Then tLap_Beton1.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "25" Then tLap_Aspal1.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "26" Then tLap_Tanah1.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "27" Then tLap_Beton2.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "28" Then tLap_Aspal2.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "29" Then tLap_Tanah2.Text = rPajak!JML_SATUAN
        
        If rPajak!KD_FASILITAS = "30" Then tLift1.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "31" Then tLift2.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "32" Then tLift3.Text = rPajak!JML_SATUAN
        
        If rPajak!KD_FASILITAS = "33" Then tTangga1.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "34" Then tTangga2.Text = rPajak!JML_SATUAN
        
        If rPajak!KD_FASILITAS = "01" Then tAC1.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "02" Then tAC2.Text = rPajak!JML_SATUAN
        
        If rPajak!KD_FASILITAS = "03" Then JPB29.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "04" Then JPB7c.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "05" Then JPB7b.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "06" Then JPB4.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "07" Then JPB5a.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "08" Then JPB5b.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "09" Then JPB13c.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "10" Then JPB13a.Text = rPajak!JML_SATUAN
        
        If rPajak!KD_FASILITAS = "11" Then cboPajak(13).Text = cboPajak(17).List(0)
        
        If rPajak!KD_FASILITAS = "43" Then JPB7d.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "44" Then tListrik.Text = rPajak!JML_SATUAN
        If rPajak!KD_FASILITAS = "45" Then JPB13d.Text = rPajak!JML_SATUAN


        rPajak.MoveNext
    Loop
     'Panggil Data Tambahan
        xJPB = Left(Trim(cboPajak(0).Text), 2)
        If xJPB = "02" Then 'Perkantoran
            iSQL = "Select * From DAT_JPB2 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "'ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                JPB29.Text = rPajak!KLS_JPB2
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "03" Then ' Pabrik
            iSQL = "Select * From DAT_JPB3 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                JPB38(0).Text = rPajak!TING_KOLOM_JPB3
                JPB38(1).Text = rPajak!LBR_BENT_JPB3
                JPB38(4).Text = rPajak!LUAS_MEZZANINE_JPB3
                JPB38(3).Text = rPajak!KELILING_DINDING_JPB3
                JPB38(2).Text = rPajak!DAYA_DUKUNG_LANTAI_JPB3
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "04" Then 'Toko/Apotik/Pasar/Ruko
            iSQL = "Select * From DAT_JPB4 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                cJPB4.Text = rPajak!KLS_JPB4
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "05" Then 'Rumah Sakit/Klinik
            iSQL = "Select * From DAT_JPB5 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                cJPB5.Text = rPajak!KLS_JPB5
                JPB5a.Text = rPajak!LUAS_KMR_JPB5_DGN_AC_SENT
                JPB5b.Text = rPajak!LUAS_RNG_LAIN_JPB5_DGN_AC_SENT
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "06" Then 'Olahraga/Rekreasi
            iSQL = "Select * From DAT_JPB6 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                cJPB6.Text = rPajak!KLS_JPB6
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "07" Then ' Hotel/Wisma
            iSQL = "Select * From DAT_JPB7 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                cJPB7a.Text = rPajak!JNS_JPB7
                cJPB7b.Text = rPajak!BINTANG_JPB7
                JPB7a.Text = rPajak!JML_KMR_JPB7
                JPB7c.Text = rPajak!LUAS_KMR_JPB7_DGN_AC_SENT
                JPB7b.Text = rPajak!LUAS_KMR_LAIN_JPB7_DGN_AC_SENT
                rPajak.MoveNext
            Loop

        ElseIf xJPB = "08" Then ' Bengkel/Gudang/Pertanian
            iSQL = "Select * From DAT_JPB8 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                JPB38(0).Text = rPajak!TING_KOLOM_JPB8
                JPB38(1).Text = rPajak!LBR_BENT_JPB8
                JPB38(4).Text = rPajak!LUAS_MEZZANINE_JPB8
                JPB38(3).Text = rPajak!KELILING_DINDING_JPB8
                JPB38(2).Text = rPajak!DAYA_DUKUNG_LANTAI_JPB8
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "09" Then
            iSQL = "Select * From DAT_JPB9 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                cJPB29.Text = rPajak!KLS_JPB9
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "12" Then 'Bangunan Parkir
             iSQL = "Select * From DAT_JPB12 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                cJPB29.Text = rPajak!TYPE_JPB12
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "13" Then ' Apartemen
            iSQL = "Select * From DAT_JPB13 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                cJPB13.Text = rPajak!KLS_JPB13
                JPB13b.Text = rPajak!JML_JPB13
                JPB13c.Text = rPajak!LUAS_JPB13_DGN_AC_SENT
                JPB13a.Text = rPajak!LUAS_JPB13_LAIN_DGN_AC_SENT
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "14" Then 'Pompa Bensin
            iSQL = "Select * From DAT_JPB14 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                txtPajak(3).Text = rPajak!LUAS_KANOPI_JPB14
                rPajak.MoveNext
            Loop
            
        ElseIf xJPB = "15" Then 'Tangki Minyak
            iSQL = "Select * From DAT_JPB15 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                cJPB15.Text = rPajak!LETAK_TANGKI_JPB15
                JPB15.Text = rPajak!KAPASITAS_TANGKI_JPB15
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "16" Then 'Gedung Sekolah
             iSQL = "Select * From DAT_JPB16 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                cJPB16.Text = rPajak!KLS_JPB16
                rPajak.MoveNext
            Loop
        ElseIf xJPB = "17" Then 'Menara
            iSQL = "Select * From DAT_JPB17 WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
            openDB (iSQL)
            If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            Do While Not rPajak.EOF
                cJPB17.Text = rPajak!TINGGI_BNG_JPB17
                cJPB17b.Text = 0
                rPajak.MoveNext
            Loop
        End If
CALL_INDIVIDU
cSubjek = "Select SUBJEK_PAJAK_ID,NOPQ from QOBJEKPAJAK WHERE NOPQ='" & Trim(aNOP.Text) & "' ORDER BY SUBJEK_PAJAK_ID"
openDB (cSubjek)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then
    txtPajak(2).Text = rPajak!SUBJEK_PAJAK_ID
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Sub CALL_INDIVIDU()
On Error GoTo Salah
StrQ1 = "Select * From DAT_NILAI_INDIVIDU WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' ORDER BY NO_BNG*1 DESC"
openDB (StrQ1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    txtPajak(0).Text = rPajak!NO_FORMULIR_INDIVIDU
    txtPajak(6).Text = rPajak!NILAI_INDIVIDU
    chPajak(3).Value = 1
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub DEL_FASILITAS1()
On Error GoTo Salah
dSQL1 = "Delete from DAT_FASILITAS_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC "
openDB (dSQL1)
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub DEL_TAMBAHAN1()
On Error GoTo Salah
Dim xJPB
xJPB = Left(cboPajak(0).Text, 2) * 1
'If (xJPB = 1 Or xJPB = 10 Or xJPB = 11 Or xJPB = 14) Or ((xJPB = 2 Or xJPB = 4 Or xJPB = 5 Or xJPB = 7 Or xJPB = 9) And xLK(2).Text <= 4) Then
If (xJPB = 1 Or xJPB = 10 Or xJPB = 11) Or ((xJPB = 2 Or xJPB = 4 Or xJPB = 5 Or xJPB = 7 Or xJPB = 9) And txtPajak(4).Text <= 4) Then
ElseIf xJPB = "02" Then 'Perkantoran
    eSQL3 = "DELETE FROM DAT_JPB2 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "03" Then ' Pabrik
    eSQL3 = "DELETE FROM DAT_JPB3 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "04" Then 'Toko/Apotik/Pasar/Ruko
    eSQL3 = "DELETE FROM DAT_JPB4 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "05" Then 'Rumah Sakit/Klinik
    eSQL3 = "DELETE FROM DAT_JPB5 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "06" Then 'Olahraga/Rekreasi
    eSQL3 = "DELETE FROM DAT_JPB6 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "07" Then ' Hotel/Wisma
    eSQL3 = "DELETE FROM DAT_JPB7 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "08" Then ' Bengkel/Gudang/Pertanian
    eSQL3 = "DELETE FROM DAT_JPB8 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "09" Then
    eSQL3 = "DELETE FROM DAT_JPB9 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "12" Then 'Bangunan Parkir
    eSQL3 = "DELETE FROM DAT_JPB12 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "13" Then ' Apartemen
    eSQL3 = "DELETE FROM DAT_JPB13  where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "14" Then 'Pompa Bensin
    eSQL3 = "DELETE FROM DAT_JPB14 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "15" Then 'Tangki Minyak
    eSQL3 = "DELETE FROM DAT_JPB15 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "16" Then 'Gedung Sekolah
    eSQL3 = "DELETE FROM DAT_JPB16 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
ElseIf xJPB = "17" Then 'Menara
    eSQL3 = "DELETE FROM DAT_JPB17 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub DEL_BANGUNAN1()
On Error Resume Next
eSQL3 = "DELETE FROM DAT_OP_BANGUNAN where (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP) =  '" & Trim(aNOP.Text) & "' AND (NO_BNG * 1 ='" & cboPajak(1).Text * 1 & "')" ' ORDER BY DAT_OP_BANGUNAN.NO_BNG DESC"
openDB (eSQL3)
End Sub
Sub DEL_INDIVIDU1()
On Error GoTo Salah
If frmObjek_Pajak_Bg.chPajak(3).Value = 1 Then
    eSQL3 = "DELETE FROM DAT_NILAI_INDIVIDU where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG*1='" & cboPajak(1).Text * 1 & "' " 'ORDER BY NO_BNG*1 DESC"
    openDB (eSQL3)
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub UP_OBJEK()
On Error GoTo Salah
totalkan_NOP


QNJOP = "SELECT * FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG ='" & xTB & "'"
openDB (QNJOP)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    If t_Nilai >= rPajak!NILAI_MIN_BNG And t_Nilai <= rPajak!NILAI_MAX_BNG Then
        t_NJOP = rPajak!NILAI_PER_M2_BNG * t_Luas * 1000
    End If
rPajak.MoveNext
Loop
    If t_NJOP = "" Or IsNull(t_NJOP) = True Then t_NJOP = 0


'iSQL4 = "UPDATE DAT_OBJEK_PAJAK SET TOTAL_LUAS_BNG = '" & t_Luas & "', NJOP_BNG = '" & t_NJOP & "' where (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "')"
'openDB (iSQL4)

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub totalkan_NOP()
On Error GoTo Salah
Dim tLuas, tNilai
iSQL3 = "SELECT * FROM DAT_OP_BANGUNAN where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' ORDER BY NO_BNG*1 DESC"
openDB (iSQL3)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If rPajak.EOF Then
    t_Luas = 0
    t_NJOP = 0
'    xxUtama = 0
'    xxMaterial = 0
'    xxFasilitas = 0
'    xxJSUSUT = 0
'    xxKSUSUT = 0
'    xxNonSusut = 0
    Exit Sub
End If
tLuas = 0: tNilai = 0
Do While Not rPajak.EOF
    tLuas = tLuas + rPajak!LUAS_BNG
    tNilai = tNilai + rPajak!NILAI_SISTEM_BNG
rPajak.MoveNext
Loop
t_Luas = tLuas
t_Nilai = tNilai / t_Luas
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub


Sub tempLog1()
On Error Resume Next
'panggil PBB terutang dari tabel SPPT
xSkg = Format(Now, "yyyy")
'xxSPPT = "select * from SPPT where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "'='" & aNOP.Text & "' and THN_PAJAK_SPPT*1='" & xSkg * 1 - 1 & "'"
xxSPPT = "select * from SPPT where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND THN_PAJAK_SPPT='" & xSkg * 1 - 1 & "'"
openDB (xxSPPT)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    ccPBB = rPajak!PBB_YG_HARUS_DIBAYAR_SPPT
    ccNJOPTKP = rPajak!NJOPTKP_SPPT
    ccKelas1 = rPajak!KD_KLS_TANAH
    ccKelas2 = rPajak!KD_KLS_BNG
rPajak.MoveNext
Loop

dNama = "select * from QObjekPajak where NOPQ='" & aNOP.Text & "' order by NOPQ asc"
openDB (dNama)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    ccNama = rPajak!Nm_wp
    ccLokasi = rPajak!JALAN_OP & ", " & rPajak!NM_KELURAHAN & " KEC. " & rPajak!NM_KECAMATAN
    rPajak.MoveNext
Loop

StrQ1 = "Select * From DAT_OBJEK_PAJAK WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' "
openDB (StrQ1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    ccLBumi = rPajak!TOTAL_LUAS_BUMI
    ccNBumi = rPajak!NJOP_BUMI
    ccNBNG = rPajak!NJOP_BNG
     ccID = rPajak!SUBJEK_PAJAK_ID
    rPajak.MoveNext
Loop

StrQ2 = "Select * From DAT_OP_BANGUNAN WHERE KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  '" & Trim(aNOP.Text) & "' AND NO_BNG='" & cboPajak(1).Text & "' ORDER BY NO_BNG*1 DESC"
openDB (StrQ2)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If rPajak.EOF Then
ccKec = Mid(Trim(aNOP.Text), 7, 3)
ccKel = Mid(Trim(aNOP.Text), 11, 3)
ccBlok = Mid(Trim(aNOP.Text), 15, 3)
ccUrut = Mid(Trim(aNOP.Text), 19, 4)
ccJenis = Right(Trim(aNOP.Text), 1)
    ccNO = 1
   ccLBangunan = 0
   zJLantai = 1
   zJenis = 0
   zTB = 0
   zTR = 0
   zKondisi = 0
   zKONSTRUKSI = 0
   zATAP = 0
   zDINDING = 0
   zLantai = 0
   zLANGIT = 0
End If

Do While Not rPajak.EOF
    ccKec = rPajak!KD_KECAMATAN
    ccKel = rPajak!KD_KELURAHAN
    ccBlok = rPajak!KD_BLOK
    ccUrut = rPajak!NO_URUT
    ccJenis = rPajak!KD_JNS_OP
    ccNO = rPajak!NO_BNG
   ccLBangunan = rPajak!LUAS_BNG
   zJLantai = rPajak!JML_LANTAI_BNG
   zJenis = rPajak!KD_JPB
   zTB = rPajak!THN_DIBANGUN_BNG
   zTR = rPajak!THN_RENOVASI_BNG
   zKondisi = rPajak!KONDISI_BNG
   zKONSTRUKSI = rPajak!JNS_KONSTRUKSI_BNG
   zATAP = rPajak!JNS_ATAP_BNG
   zDINDING = rPajak!KD_DINDING
   zLantai = rPajak!KD_LANTAI
   zLANGIT = rPajak!KD_LANGIT_LANGIT
   'zListrik = rPajak!LISTRIK
   
    
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
    rPajak!KD_KECAMATAN = ccKec
    rPajak!KD_KELURAHAN = ccKel
    rPajak!KD_BLOK = ccBlok
    rPajak!NO_URUT = ccUrut
    rPajak!KD_JNS_OP = ccJenis
    rPajak!NO_BNG = ccNO
    rPajak!SUBJEK_PAJAK_ID = ccID
   rPajak!NO_FORMULIR_SPOP = txtPajak(0).Text 'Formulir/Dokumen
   rPajak!NO_PERSIL = "-" 'tBumi(10).Text 'Persil
   rPajak!JALAN_OP = ccLokasi
   rPajak!BLOK_KAV_NO_OP = "-" 'tBumi(11).Text 'Blok/Kav
   rPajak!RW_OP = "-" 'tBumi(8).Text 'RW
   rPajak!RT_OP = "-" 'tBumi(9).Text 'RT
   
   rPajak!KD_STATUS_WP = "2" 'Left(Trim(cboStatus.Text), 1)
   rPajak!TOTAL_LUAS_BUMI = ccLBumi
    rPajak!NJOP_BUMI = ccNBumi
    
    rPajak!PBB_Terutang = ccPBB
    rPajak!KD_STATUS_CABANG = 0
   rPajak!TOTAL_LUAS_BNG = ccLBangunan 'txtPajak(3).Text ' xxTotalBng
   rPajak!NJOP_BNG = ccNBNG
   rPajak!STATUS_PETA_OP = "1"
   If chPajak(2).Value = 1 Then
        rPajak!JNS_TRANSAKSI_OP = 2
   Else
        rPajak!JNS_TRANSAKSI_OP = 3
   End If
    rPajak!TGL_PENDATAAN_OP = Format(dtPajak(0).Value, "dd/mm/yyyy") 'Tanggal Pendataan
   rPajak!TGL_PEMERIKSAAN_OP = Format(dtPajak(1).Value, "dd/mm/yyyy") 'Tanggal Pemeriksaan
   'rPajak!TGL_PEREKAMAN_OP = Format(dtPajak(2).Value, "dd/mm/yyyy") & Format(Now, "hh:mm:ss")  'Tanggal Perekaman
   rPajak!TGL_PEREKAMAN_OP = Format(Now, "dd/mm/yyyy hh:mm:ss") 'Tanggal Perekaman
   rPajak!NIP_PENDATA = txtPajak(7).Text 'NIP Pendata
   rPajak!NIP_PEMERIKSA_OP = txtPajak(8).Text 'NIP Pemeriksa
   rPajak!NIP_PEREKAM_OP = txtPajak(9).Text 'NIP Perekam
   rPajak!JLANTAI = zJLantai
   rPajak!JENIS_BNG = zJenis
   rPajak!THN_DIBANGUN = zTB
   rPajak!THN_DIRENOVASI = zTR
   rPajak!KONDISI = zKondisi
   rPajak!KONSTRUKSI = zKONSTRUKSI
   rPajak!ATAP = zATAP
   rPajak!DINDING = zDINDING
   rPajak!Lantai = zLantai
   rPajak!LANGIT = zLANGIT
   rPajak!LISTRIK = zListik
   rPajak!Nm_wp = ccNama
   If ccNJOPTKP = Empty Then
    rPajak!NJOPTKP = 0
   Else
    rPajak!NJOPTKP = ccNJOPTKP
   End If
   rPajak!KD_KLS_TANAH = ccKelas1
   rPajak!KD_KLS_BNG = ccKelas2
    rPajak.Update
End Sub

'Log untuk penghapusan data Bangunan

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
    cTotal2 = rPajak!TOTAL_LUAS_BNG
    cBNG = rPajak!NJOP_BNG
    ccNO = rPajak!NO_BNG
    cNama = rPajak!Nm_wp
    cNJOPTKP = rPajak!NJOPTKP
    rPajak!JLANTAI = txtPajak(4).Text
   rPajak!JENIS_BNG = cboPajak(0).Text
   rPajak!THN_DIBANGUN = cboPajak(2).Text
   rPajak!THN_DIRENOVASI = cboPajak(3).Text
   rPajak!KONDISI = cboPajak(4).Text
   rPajak!KONSTRUKSI = cboPajak(5).Text
   rPajak!ATAP = cboPajak(6).Text
   rPajak!DINDING = cboPajak(7).Text
   rPajak!Lantai = cboPajak(8).Text
   rPajak!LANGIT = cboPajak(9).Text
   rPajak!LISTRIK = tListrik.Text
   cPBB = rPajak!PBB_Terutang
   cKelas1 = rPajak!KD_KLS_TANAH
   cKelas2 = rPajak!KD_KLS_BNG
rPajak.MoveNext
Loop
     upSTR = "Select * From LogUtama"
    openDB (upSTR)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    rPajak.AddNew

    'Data Lama
     rPajak!NOP = "12.12." & cKec & "." & cKel & "." & cblok & "-" & cUrut & "." & cJenis
     rPajak!SUBJEK_PAJAK_ID = cID
     rPajak!Lokasi = cLokasi
     rPajak!KD_STATUS_WP = cWP
     rPajak!TOTAL_LUAS_BUMI = cTotal1
     rPajak!NJOP_BUMI = cBumi
      rPajak!TOTAL_LUAS_BNG = cTotal2
     rPajak!NJOP_BNG = cBNG
     rPajak!NO_BNG = ccNO
     rPajak!Nm_wp = cNama
     rPajak!NM_WP1 = cNama
    'Data Baru
    
    rPajak!NOP1 = aNOP.Text
    rPajak!subjek_pajak_id1 = "-"
    rPajak!lokasi1 = cLokasi
   rPajak!kd_status_wp1 = "-"
   rPajak!TOTAL_LUAS_BUMI1 = cTotal1
   rPajak!NJOP_BUMI1 = cBumi
   If chPajak(1).Value = 1 Then
        rPajak!JNS_TRANSAKSI_OP = 2
   ElseIf chPajak(2).Value = 1 Then
        rPajak!JNS_TRANSAKSI_OP = 3
   End If
    'rPajak!TGL_PEREKAMAN_OP = Format(dtPajak(2).Value, "dd/mm/yyyy") & Format(Now, "hh:mm:ss") 'Tanggal Perekaman
       rPajak!TGL_PEREKAMAN_OP = Format(Now, "dd/mm/yyyy hh:mm:ss")  'Tanggal Perekaman
   rPajak!NIP_PEREKAM_OP = txtPajak(9).Text 'NIP Perekam
   rPajak!TOTAL_LUAS_BNG = 0
     
   rPajak!NJOP_BNG = 0
   
   rPajak!PBB_Terutang = cPBB
   rPajak!NJOPTKP = cNJOPTKP
   rPajak!NJOPTKP1 = 0
   rPajak!TOTAL_LUAS_BNG1 = 0
    rPajak!NJOP_BNG1 = 0 'xxNJOPBng
   rPajak!PBB_TERUTANG1 = 0
   
   
' If zJLantai * 1 = txtPajak(4).Text * 1 Then rPajak!JLANTAI = "-" Else rPajak!JLANTAI = txtPajak(4).Text
'   If zJenis * 1 = Left(cboPajak(0).Text, 2) * 1 Then rPajak!JENIS_BNG = "-" Else rPajak!JENIS_BNG = cboPajak(0).Text
'   If zTB * 1 = cboPajak(2).Text * 1 Then rPajak!THN_DIBANGUN = "-" Else rPajak!THN_DIBANGUN = cboPajak(2).Text
'   If zTR * 1 = cboPajak(3).Text Then rPajak!THN_DIRENOVASI = "-" Else rPajak!THN_DIRENOVASI = cboPajak(3).Text
'   If zKondisi * 1 = Left(cboPajak(4).Text, 2) * 1 Then rPajak!KONDISI = "-" Else rPajak!KONDISI = cboPajak(4).Text
'   If zKONSTRUKSI * 1 = Left(cboPajak(5).Text, 2) * 1 Then rPajak!KONSTRUKSI = "-" Else rPajak!KONSTRUKSI = cboPajak(5).Text
'   If zATAP = Left(cboPajak(6).Text, 1) Then rPajak!ATAP = "-" Else rPajak!ATAP = cboPajak(6).Text
'   If zDINDING = Left(cboPajak(7).Text, 1) Then rPajak!DINDING = "-" Else rPajak!DINDING = cboPajak(7).Text
'   If zLantai = Left(cboPajak(8).Text, 1) Then rPajak!Lantai = "-" Else rPajak!Lantai = cboPajak(8).Text
'   If zLANGIT = Left(cboPajak(9).Text, 1) Then rPajak!LANGIT = "-" Else rPajak!LANGIT = cboPajak(9).Text
    rPajak!JLANTAI = txtPajak(4).Text
   rPajak!JENIS_BNG = cboPajak(0).Text
   rPajak!THN_DIBANGUN = cboPajak(2).Text
   rPajak!THN_DIRENOVASI = cboPajak(3).Text
   rPajak!KONDISI = cboPajak(4).Text
   rPajak!KONSTRUKSI = cboPajak(5).Text
   rPajak!ATAP = cboPajak(6).Text
   rPajak!DINDING = cboPajak(7).Text
   rPajak!Lantai = cboPajak(8).Text
   rPajak!LANGIT = cboPajak(9).Text
   'If zLISTRIK * 1 = tListrik.Text * 1 Then rPajak!LISTRIK = 0 Else rPajak!LISTRIK = tListrik.Text
   rPajak!KD_KLS_TANAH = cKelas1
   rPajak!KD_KLS_BNG = cKelas2
   rPajak!KET = tKet.Text & " (Hapus Data Bangunan)"
   rPajak!xFlag = "1"
    rPajak.Update
'   strFlag = "Select * from logutama where NOP='" & aNOP.Text & "' order by NOP asc"
'   openDB (strFlag)
'   If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'   Do While Not rPajak.EOF
'    rPajak!xFlag = "2"
'    rPajak.MoveNext
'   Loop
End Sub
