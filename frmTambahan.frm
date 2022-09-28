VERSION 5.00
Begin VB.Form frmTambahan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Bangunan"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7455
   Icon            =   "frmTambahan.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   7455
   Begin VB.Frame FF1 
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
      Left            =   165
      TabIndex        =   62
      Top             =   60
      Width           =   7260
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
         TabIndex        =   1
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
         Index           =   4
         Left            =   5535
         TabIndex        =   4
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
         Index           =   2
         Left            =   2310
         TabIndex        =   2
         Top             =   810
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
         Left            =   2295
         TabIndex        =   0
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
         Index           =   3
         Left            =   5535
         TabIndex        =   3
         Top             =   180
         Width           =   1650
      End
      Begin VB.Label Label28 
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
         TabIndex        =   67
         Top             =   840
         Width           =   2205
      End
      Begin VB.Label Label36 
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
         TabIndex        =   66
         Top             =   195
         Width           =   1470
      End
      Begin VB.Label Label37 
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
         TabIndex        =   65
         Top             =   525
         Width           =   1500
      End
      Begin VB.Label Label38 
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
         Left            =   150
         TabIndex        =   64
         Top             =   255
         Width           =   1770
      End
      Begin VB.Label Label40 
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
         TabIndex        =   63
         Top             =   555
         Width           =   1470
      End
   End
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
      Height          =   7950
      Left            =   60
      TabIndex        =   33
      Top             =   1440
      Width           =   7455
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
         Left            =   105
         TabIndex        =   70
         Top             =   4980
         Width           =   7260
         Begin VB.TextBox JPB17B 
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
            Left            =   5040
            TabIndex        =   21
            Top             =   225
            Width           =   2040
         End
         Begin VB.TextBox JPB17 
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
            Left            =   1695
            TabIndex        =   20
            Top             =   210
            Width           =   2040
         End
         Begin VB.Label Label4 
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
            Left            =   4005
            TabIndex        =   72
            Top             =   270
            Width           =   1770
         End
         Begin VB.Label Label3 
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
            TabIndex        =   71
            Top             =   255
            Width           =   1770
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
         Left            =   0
         TabIndex        =   60
         Top             =   105
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
            Left            =   5625
            TabIndex        =   6
            Top             =   210
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
            Left            =   1425
            TabIndex        =   5
            Top             =   225
            Width           =   2670
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
            Left            =   4200
            TabIndex        =   73
            Top             =   150
            Width           =   1485
         End
         Begin VB.Label Label54 
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
            TabIndex        =   61
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
         Height          =   1470
         Index           =   4
         Left            =   90
         TabIndex        =   54
         Top             =   3180
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
            TabIndex        =   18
            Top             =   990
            Width           =   1545
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
            Left            =   5475
            TabIndex        =   16
            Top             =   270
            Width           =   1560
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
            Height          =   300
            Left            =   1635
            TabIndex        =   13
            Top             =   225
            Width           =   2040
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
            TabIndex        =   17
            Top             =   630
            Width           =   1545
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
            Height          =   300
            Left            =   1635
            TabIndex        =   14
            Top             =   585
            Width           =   2025
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
            TabIndex        =   15
            Top             =   960
            Width           =   2025
         End
         Begin VB.Label Label1 
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
            Left            =   3885
            TabIndex        =   68
            Top             =   975
            Width           =   1500
         End
         Begin VB.Label Label60 
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
            TabIndex        =   59
            Top             =   960
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
            TabIndex        =   58
            Top             =   165
            Width           =   1485
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
            TabIndex        =   57
            Top             =   585
            Width           =   1500
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
            TabIndex        =   56
            Top             =   255
            Width           =   1770
         End
         Begin VB.Label Label59 
            Caption         =   "Kelas Hotel"
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
            TabIndex        =   55
            Top             =   585
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
         TabIndex        =   52
         Top             =   4665
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
            Left            =   1650
            TabIndex        =   19
            Top             =   225
            Width           =   5550
         End
         Begin VB.Label Label35 
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
            TabIndex        =   53
            Top             =   255
            Width           =   1770
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
         Height          =   1425
         Index           =   6
         Left            =   165
         TabIndex        =   47
         Top             =   5385
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
            TabIndex        =   26
            Top             =   1005
            Width           =   1860
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
            TabIndex        =   23
            Top             =   615
            Width           =   2025
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
            TabIndex        =   25
            Top             =   615
            Width           =   1845
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
            Height          =   300
            Left            =   1680
            TabIndex        =   22
            Top             =   225
            Width           =   2040
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
            TabIndex        =   24
            Top             =   210
            Width           =   1830
         End
         Begin VB.Label Label2 
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
            TabIndex        =   69
            Top             =   945
            Width           =   2010
         End
         Begin VB.Label Label39 
            Caption         =   "Jumlah Apartemen"
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
            Left            =   3885
            TabIndex        =   51
            Top             =   240
            Width           =   1770
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
            TabIndex        =   50
            Top             =   270
            Width           =   1770
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
            TabIndex        =   49
            Top             =   555
            Width           =   1500
         End
         Begin VB.Label Label52 
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
            TabIndex        =   48
            Top             =   555
            Width           =   1485
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
         Left            =   0
         TabIndex        =   44
         Top             =   6825
         Width           =   7260
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
            Height          =   300
            ItemData        =   "frmTambahan.frx":0CCA
            Left            =   4680
            List            =   "frmTambahan.frx":0CD4
            TabIndex        =   28
            Top             =   195
            Width           =   2520
         End
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
            TabIndex        =   27
            Top             =   195
            Width           =   1845
         End
         Begin VB.Label Label62 
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
            TabIndex        =   46
            Top             =   255
            Width           =   960
         End
         Begin VB.Label Label63 
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
            TabIndex        =   45
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
         TabIndex        =   40
         Top             =   1575
         Width           =   7260
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
            TabIndex        =   10
            Top             =   540
            Width           =   1545
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
            TabIndex        =   9
            Top             =   225
            Width           =   4890
         End
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
            Left            =   5580
            TabIndex        =   11
            Top             =   555
            Width           =   1650
         End
         Begin VB.Label Label65 
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
            TabIndex        =   43
            Top             =   255
            Width           =   1770
         End
         Begin VB.Label Label64 
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
            TabIndex        =   42
            Top             =   525
            Width           =   1500
         End
         Begin VB.Label Label61 
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
            TabIndex        =   41
            Top             =   525
            Width           =   1485
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
         Height          =   660
         Index           =   1
         Left            =   90
         TabIndex        =   38
         Top             =   660
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
            TabIndex        =   8
            Top             =   225
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
            Left            =   1335
            TabIndex        =   7
            Top             =   225
            Width           =   2670
         End
         Begin VB.Label Label5 
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
            Left            =   4080
            TabIndex        =   74
            Top             =   135
            Width           =   1485
         End
         Begin VB.Label Label53 
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
            TabIndex        =   39
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
         TabIndex        =   36
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
            Height          =   300
            Left            =   2295
            TabIndex        =   12
            Top             =   225
            Width           =   4890
         End
         Begin VB.Label Label55 
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
            TabIndex        =   37
            Top             =   255
            Width           =   1770
         End
      End
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
         TabIndex        =   34
         Top             =   7215
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
            Height          =   300
            Left            =   2295
            TabIndex        =   29
            Top             =   225
            Width           =   4890
         End
         Begin VB.Label Label66 
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
            TabIndex        =   35
            Top             =   255
            Width           =   1770
         End
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
      Height          =   435
      Left            =   4350
      TabIndex        =   32
      Top             =   9510
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3375
      TabIndex        =   31
      Top             =   9510
      Width           =   990
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   2400
      TabIndex        =   30
      Top             =   9510
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   9195
      Left            =   0
      Picture         =   "frmTambahan.frx":0CE9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "frmTambahan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cJPB12_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select

End Sub

Private Sub cJPB12_LostFocus()
On Error Resume Next
If cJPB12.Text = "" Then cJPB12.Text = cJPB12.List(0)
cJPB12.Text = cJPB12.List(cJPB12.Text - 1)
    If cJPB12.Text <> cJPB12.List(0) And cJPB12.Text <> cJPB12.List(1) And cJPB12.Text <> cJPB12.List(2) And cJPB12.Text <> cJPB12.List(3) Then
        cJPB12.Text = cJPB12.List(0)
    End If
End Sub

Private Sub cJPB13_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select

End Sub

Private Sub cJPB13_LostFocus()
On Error Resume Next
If cJPB13.Text = "" Then cJPB13.Text = cJPB13.List(0)
cJPB13.Text = cJPB13.List(cJPB13.Text - 1)
    If cJPB13.Text <> cJPB13.List(0) And cJPB13.Text <> cJPB13.List(1) And cJPB13.Text <> cJPB13.List(2) And cJPB13.Text <> cJPB13.List(3) Then
        cJPB13.Text = cJPB13.List(0)
    End If
End Sub

Private Sub cJPB15_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select

End Sub

Private Sub cJPB15_LostFocus()
On Error Resume Next
If cJPB15.Text = "" Then cJPB15.Text = cJPB15.List(0)
cJPB15.Text = cJPB15.List(cJPB15.Text - 1)
    If cJPB15.Text <> cJPB15.List(0) And cJPB15.Text <> cJPB15.List(1) Then
        cJPB15.Text = cJPB15.List(0)
    End If
End Sub

Private Sub cJPB16_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select

End Sub

Private Sub cJPB16_LostFocus()
On Error Resume Next
If cJPB16.Text = "" Then cJPB16.Text = cJPB16.List(0)
cJPB16.Text = cJPB16.List(cJPB16.Text - 1)
    If cJPB16.Text <> cJPB16.List(0) And cJPB16.Text <> cJPB16.List(1) Then
        cJPB16.Text = cJPB16.List(0)
    End If
End Sub

Private Sub cJPB29_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub cJPB29_LostFocus()
On Error Resume Next
'If cJPB29.Text = "" Then cJPB29.Text = cJPB29.List(0)

        
    cJPB29.Text = cJPB29.List(cJPB29.Text - 1)
    If cJPB29.Text <> cJPB29.List(0) And cJPB29.Text <> cJPB29.List(1) And cJPB29.Text <> cJPB29.List(2) And cJPB29.Text <> cJPB29.List(3) Then
        cJPB29.Text = cJPB29.List(0)
    End If

End Sub

Private Sub cJPB4_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If

End Sub

Private Sub cJPB4_LostFocus()
On Error Resume Next
If cJPB4.Text = "" Then cJPB4.Text = cJPB4.List(0)
cJPB4.Text = cJPB4.List(cJPB4.Text - 1)
    If cJPB4.Text <> cJPB4.List(0) And cJPB4.Text <> cJPB4.List(1) And cJPB4.Text <> cJPB4.List(2) Then
        cJPB4.Text = cJPB4.List(0)
    End If
End Sub

Private Sub cJPB5_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If

End Sub

Private Sub cJPB5_LostFocus()
On Error Resume Next
If cJPB5.Text = "" Then cJPB5.Text = cJPB5.List(0)
 cJPB5.Text = cJPB5.List(cJPB5.Text - 1)
    If cJPB5.Text <> cJPB5.List(0) And cJPB5.Text <> cJPB5.List(1) And cJPB5.Text <> cJPB5.List(2) And cJPB5.Text <> cJPB5.List(3) Then
        cJPB5.Text = cJPB5.List(0)
    End If
End Sub

Private Sub cJPB6_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If

End Sub

Private Sub cJPB6_LostFocus()
On Error Resume Next
If cJPB6.Text = "" Then cJPB6.Text = cJPB6.List(0)
cJPB6.Text = cJPB6.List(cJPB6.Text - 1)
    If cJPB6.Text <> cJPB6.List(0) And cJPB6.Text <> cJPB6.List(1) Then
        cJPB6.Text = cJPB6.List(0)
    End If
End Sub

Private Sub cJPB7a_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If

End Sub

Private Sub cJPB7a_LostFocus()
On Error Resume Next
'If cJPB7a.Text = "" Then cJPB7a.Text = cJPB7a.List(0)
cJPB7a.Text = cJPB7a.List(cJPB7a.Text - 1)
    If cJPB7a.Text <> cJPB7a.List(0) And cJPB7a.Text <> cJPB7a.List(1) Then
        cJPB7a.Text = cJPB7a.List(0)
    End If
End Sub

Private Sub cJPB7b_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If

End Sub

Private Sub cJPB7b_LostFocus()
On Error Resume Next
If cJPB7b.Text = "" Then cJPB7b.Text = cJPB7b.List(0)
cJPB7b.Text = cJPB7b.List(cJPB7b.Text)
    If cJPB7b.Text <> cJPB7b.List(0) And cJPB7b.Text <> cJPB7b.List(1) And cJPB7b.Text <> cJPB7b.List(2) And cJPB7b.Text <> cJPB7b.List(3) And cJPB7b.Text <> cJPB7b.List(4) And cJPB7b.Text <> cJPB7b.List(5) Then
        cJPB7b.Text = cJPB7b.List(0)
    End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = 0
        Control.Alignment = 1
    End If
    If TypeOf Control Is ComboBox Then
        Control.Text = Control.List(0)
    End If
Next

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo Salah
For i = 0 To 4
    frmObjek_Pajak_Bg.JPB38(i).Text = 0
Next
frmObjek_Pajak_Bg.cJPB17.Text = 0: frmObjek_Pajak_Bg.cJPB17b.Text = 0: frmObjek_Pajak_Bg.JPB29.Text = 0: frmObjek_Pajak_Bg.JPB4.Text = 0
frmObjek_Pajak_Bg.JPB5a.Text = 0: frmObjek_Pajak_Bg.JPB5b.Text = 0
frmObjek_Pajak_Bg.JPB7a.Text = 0: frmObjek_Pajak_Bg.JPB7b.Text = 0: frmObjek_Pajak_Bg.JPB7c.Text = 0: frmObjek_Pajak_Bg.JPB7d.Text = 0
frmObjek_Pajak_Bg.JPB13a.Text = 0: frmObjek_Pajak_Bg.JPB13b.Text = 0: frmObjek_Pajak_Bg.JPB13c.Text = 0: frmObjek_Pajak_Bg.JPB13d.Text = 0
frmObjek_Pajak_Bg.JPB15.Text = 0
Select Case xxNon
Case 2, 9
    frmObjek_Pajak_Bg.cJPB29.Text = Left(Trim(cJPB29.Text), 1)
    frmObjek_Pajak_Bg.JPB29.Text = Trim(JPB29.Text)
Case 3, 8
    frmObjek_Pajak_Bg.JPB38(0).Text = JPB38(0).Text
    frmObjek_Pajak_Bg.JPB38(1).Text = JPB38(1).Text
    frmObjek_Pajak_Bg.JPB38(2).Text = JPB38(2).Text
    frmObjek_Pajak_Bg.JPB38(3).Text = JPB38(3).Text
    frmObjek_Pajak_Bg.JPB38(4).Text = JPB38(4).Text
Case 4
    frmObjek_Pajak_Bg.cJPB4.Text = Left(Trim(cJPB4.Text), 1)
    frmObjek_Pajak_Bg.JPB4.Text = Trim(JPB4.Text)
Case 5
    frmObjek_Pajak_Bg.cJPB5.Text = Left(Trim(cJPB5.Text), 1)
    frmObjek_Pajak_Bg.JPB5a.Text = JPB5a.Text
    frmObjek_Pajak_Bg.JPB5b.Text = JPB5b.Text
Case 6
    frmObjek_Pajak_Bg.cJPB6.Text = Left(Trim(cJPB6.Text), 1)
Case 7
    frmObjek_Pajak_Bg.cJPB7a.Text = Left(Trim(cJPB7a.Text), 1)
    frmObjek_Pajak_Bg.cJPB7b.Text = Left(Trim(cJPB7b.Text), 1) * 1
    frmObjek_Pajak_Bg.JPB7a.Text = JPB7a.Text
    frmObjek_Pajak_Bg.JPB7b.Text = JPB7b.Text
    frmObjek_Pajak_Bg.JPB7c.Text = JPB7c.Text
    frmObjek_Pajak_Bg.JPB7d.Text = JPB7d.Text
Case 12
    frmObjek_Pajak_Bg.cJPB12.Text = Left(Trim(cJPB12.Text), 1)
Case 13
    frmObjek_Pajak_Bg.cJPB13.Text = Left(Trim(cJPB13.Text), 1)
    frmObjek_Pajak_Bg.JPB13a.Text = JPB13a.Text
    frmObjek_Pajak_Bg.JPB13b.Text = JPB13b.Text
    frmObjek_Pajak_Bg.JPB13c.Text = JPB13c.Text
    frmObjek_Pajak_Bg.JPB13d.Text = JPB13d.Text
Case 15
    frmObjek_Pajak_Bg.cJPB15.Text = Left(Trim(cJPB15.Text), 1)
    frmObjek_Pajak_Bg.JPB15.Text = JPB15.Text
Case 16
    frmObjek_Pajak_Bg.cJPB16.Text = Left(Trim(cJPB16.Text), 1)
Case 17
    frmObjek_Pajak_Bg.cJPB17.Text = JPB17.Text
    frmObjek_Pajak_Bg.cJPB17b.Text = JPB17B.Text
End Select
Unload Me
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub Form_Activate()
On Error Resume Next
Screen.MousePointer = vbHourglass
Me.Width = 7700
Me.Height = 1815 ' 10620
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
'xxNon = InputBox("Data Tambangan [1-16]", "Input")
For i = 0 To 8
    fDT(i).Top = 100
    fDT(i).Visible = False
Next
FF2.Top = 10
Select Case xxNon
    'Case 1, 10, 11, 14
     '       End
    Case 2, 9
        FF1.Visible = False
        
        fDT(0).Visible = True
        Me.Caption = fDT(0).Caption
        fDT(0).Caption = ""
        FF2.Height = fDT(0).Height + 500
        Me.Height = FF2.Height + 1200
        cmdSave.Top = FF2.Height + 70
        cmdCancel.Top = FF2.Height + 70
        cmdExit.Top = FF2.Height + 70
        cJPB29.SetFocus
    Case 3, 8
        FF2.Visible = False
        Me.Caption = "[DATA TAMBAHAN UNTUK JPB = 3/8] (Pabrik/Bengkel/Gudang/Pertanian)"
        Me.Height = FF1.Height + 1200
        cmdSave.Top = FF1.Height + 120
        cmdCancel.Top = FF1.Height + 120
        cmdExit.Top = FF1.Height + 120
        JPB38(0).SetFocus
    Case 4
        FF1.Visible = False
        fDT(1).Visible = True
        Me.Caption = fDT(1).Caption
        fDT(1).Caption = ""
        FF2.Height = fDT(1).Height + 500
        Me.Height = FF2.Height + 1200
        cmdSave.Top = FF2.Height + 100
        cmdCancel.Top = FF2.Height + 100
        cmdExit.Top = FF2.Height + 100
        cJPB4.SetFocus
    Case 5
          FF1.Visible = False
        fDT(2).Visible = True
        Me.Caption = fDT(2).Caption
        fDT(2).Caption = ""
        FF2.Height = fDT(2).Height + 500
        Me.Height = FF2.Height + 1200
        cmdSave.Top = FF2.Height + 100
        cmdCancel.Top = FF2.Height + 100
        cmdExit.Top = FF2.Height + 100
        cJPB5.SetFocus
    Case 6
        FF1.Visible = False
        fDT(3).Visible = True
        Me.Caption = fDT(3).Caption
        fDT(3).Caption = ""
        FF2.Height = fDT(3).Height + 500
        Me.Height = FF2.Height + 1200
        cmdSave.Top = FF2.Height + 100
        cmdCancel.Top = FF2.Height + 100
        cmdExit.Top = FF2.Height + 100
        cJPB6.SetFocus
    Case 7
        FF1.Visible = False
        fDT(4).Visible = True
        Me.Caption = fDT(4).Caption
        fDT(4).Caption = ""
        FF2.Height = fDT(4).Height + 500
        Me.Height = FF2.Height + 1200
        cmdSave.Top = FF2.Height + 100
        cmdCancel.Top = FF2.Height + 100
        cmdExit.Top = FF2.Height + 100
        cJPB7a.SetFocus
    Case 12
        FF1.Visible = False
        fDT(5).Visible = True
        Me.Caption = fDT(5).Caption
        fDT(5).Caption = ""
        FF2.Height = fDT(5).Height + 500
        Me.Height = FF2.Height + 1200
        cmdSave.Top = FF2.Height + 100
        cmdCancel.Top = FF2.Height + 100
        cmdExit.Top = FF2.Height + 100
        cJPB12.SetFocus
    Case 13
        FF1.Visible = False
        fDT(6).Visible = True
        Me.Caption = fDT(6).Caption
        fDT(6).Caption = ""
        FF2.Height = fDT(6).Height + 500
        Me.Height = FF2.Height + 1200
        cmdSave.Top = FF2.Height + 100
        cmdCancel.Top = FF2.Height + 100
        cmdExit.Top = FF2.Height + 100
        cJPB13.SetFocus
    Case 15
        FF1.Visible = False
        fDT(7).Visible = True
        Me.Caption = fDT(7).Caption
        fDT(7).Caption = ""
        FF2.Height = fDT(7).Height + 500
        Me.Height = FF2.Height + 1200
        cmdSave.Top = FF2.Height + 100
        cmdCancel.Top = FF2.Height + 100
        cmdExit.Top = FF2.Height + 100
        'JPB15.SetFocus
    Case 16
        FF1.Visible = False
        fDT(8).Visible = True
        Me.Caption = fDT(8).Caption
        fDT(8).Caption = ""
        FF2.Height = fDT(8).Height + 500
        Me.Height = FF2.Height + 1200
        cmdSave.Top = FF2.Height + 100
        cmdCancel.Top = FF2.Height + 100
        cmdExit.Top = FF2.Height + 100
        cJPB16.SetFocus
    Case 17
        FF1.Visible = False
        fDT(9).Visible = True
        fDT(9).Top = 80
        Me.Caption = fDT(9).Caption
        fDT(9).Caption = ""
        FF2.Height = fDT(9).Height + 500
        Me.Height = FF2.Height + 1200
        cmdSave.Top = FF2.Height + 100
        cmdCancel.Top = FF2.Height + 100
        cmdExit.Top = FF2.Height + 100
        JPB17.SetFocus
End Select
cJPB29.Clear
cJPB29.Text = "1-KELAS 1"
cJPB29.AddItem "1-KELAS 1"
cJPB29.AddItem "2-KELAS 2"
cJPB29.AddItem "3-KELAS 3"
cJPB29.AddItem "4-KELAS 4"
cJPB4.Clear
cJPB4.Text = "1-KELAS 1"
cJPB4.AddItem "1-KELAS 1"
cJPB4.AddItem "2-KELAS 2"
cJPB4.AddItem "3-KELAS 3"
cJPB5.Clear
cJPB5.Text = "1-KELAS 1"
cJPB5.AddItem "1-KELAS 1"
cJPB5.AddItem "2-KELAS 2"
cJPB5.AddItem "3-KELAS 3"
cJPB5.AddItem "4-KELAS 4"
cJPB6.Clear
cJPB6.Text = "1-KELAS 1"
cJPB6.AddItem "1-KELAS 1"
cJPB6.AddItem "2-KELAS 2"
cJPB7a.Clear
cJPB7a.Text = "1-NON RESORT"
cJPB7a.AddItem "1-NON RESORT"
cJPB7a.AddItem "2-RESORT"
cJPB7b.Clear
cJPB7b.Text = "0-NON BINTANG"
cJPB7b.AddItem "0-NON BINTANG"
cJPB7b.AddItem "1-BINTANG 5"
cJPB7b.AddItem "2-BINTANG 4"
cJPB7b.AddItem "3-BINTANG 3"
cJPB7b.AddItem "4-BINTANG 2"
cJPB7b.AddItem "5-BINTANG 1"
cJPB12.Clear
cJPB12.Text = "1-TIPE 1"
cJPB12.AddItem "1-TIPE 1"
cJPB12.AddItem "2-TIPE 2"
cJPB12.AddItem "3-TIPE 3"
cJPB12.AddItem "4-TIPE 4"
cJPB13.Clear
cJPB13.Text = "1-KELAS 1"
cJPB13.AddItem "1-KELAS 1"
cJPB13.AddItem "2-KELAS 2"
cJPB13.AddItem "3-KELAS 3"
cJPB13.AddItem "4-KELAS 4"
cJPB15.Clear
cJPB15.Text = "1-DIATAS TANAH"
cJPB15.AddItem "1-DIATAS TANAH"
cJPB15.AddItem "2-DIBAWAH TANAH"
cJPB16.Clear
cJPB16.Text = "1-KELAS 1"
cJPB16.AddItem "1-KELAS 1"
cJPB16.AddItem "2-KELAS 2"
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = 0
    End If
Next
cmdCancel_Click
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
xxJPB = frmObjek_Pajak_Bg.cboPajak(0).Text
End Sub


Private Sub JPB13a_Change()
On Error Resume Next
JPB13d.Text = JPB13a.Text
End Sub

Private Sub JPB13a_GotFocus()
On Error Resume Next
Call c_blok(JPB13a)
End Sub

Private Sub JPB13a_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select

End Sub

Private Sub JPB13a_LostFocus()
On Error Resume Next
Call c_Kosong(JPB13a)
End Sub

Private Sub JPB13b_GotFocus()
On Error Resume Next
Call c_blok(JPB13b)
End Sub

Private Sub JPB13b_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select

End Sub

Private Sub JPB13b_LostFocus()
On Error Resume Next
Call c_Kosong(JPB13b)
End Sub

Private Sub JPB13c_GotFocus()
On Error Resume Next
Call c_blok(JPB13c)
End Sub

Private Sub JPB13c_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select

End Sub

Private Sub JPB13c_LostFocus()
On Error Resume Next
Call c_Kosong(JPB13c)
End Sub

Private Sub JPB13d_GotFocus()
On Error Resume Next
Call c_blok(JPB13d)
End Sub

Private Sub JPB13d_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select

End Sub

Private Sub JPB13d_LostFocus()
On Error Resume Next
Call c_Kosong(JPB13d)
End Sub

Private Sub JPB15_GotFocus()
On Error Resume Next
Call c_blok(JPB15)
End Sub

Private Sub JPB15_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select

End Sub

Private Sub JPB15_LostFocus()
On Error Resume Next
Call c_Kosong(JPB15)
End Sub

Private Sub JPB17_GotFocus()
On Error Resume Next
Call c_blok(JPB17)
JPB17.Alignment = 0
End Sub

Private Sub JPB17_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select

End Sub

Private Sub JPB17_LostFocus()
On Error Resume Next
If JPB17.Text = "" Or JPB17.Text = "." Or JPB17.Text = "," Or JPB17.Text = "," Then JPB17.Text = 0
JPB17.Alignment = 1
End Sub

Private Sub JPB17B_GotFocus()
On Error Resume Next
Call c_blok(JPB17B)
JPB17B.Alignment = 0
End Sub

Private Sub JPB17B_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub JPB17B_LostFocus()
On Error Resume Next
If JPB17B.Text = "" Or JPB17B.Text = "." Or JPB17B.Text = "," Or JPB17B.Text = "," Then JPB17B.Text = 0
JPB17B.Alignment = 1
End Sub

Private Sub JPB29_GotFocus()
On Error Resume Next
Call c_blok(JPB29)
JPB29.Alignment = 0
End Sub

Private Sub JPB29_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789,.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
End If
End Sub

Private Sub JPB29_LostFocus()
On Error Resume Next
Call c_Kosong(JPB29)
JPB29.Alignment = 1

End Sub

Private Sub JPB38_GotFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 0 To 4
    Call c_blok(JPB38(Index))
    JPB38(Index).Alignment = 0
End Select
End Sub

Private Sub JPB38_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
Select Case Index
Case 0 To 4
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Select
End Sub

Private Sub JPB38_LostFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 0 To 4
    Call c_Kosong(JPB38(Index))
    JPB38(Index).Alignment = 1
End Select
End Sub

Private Sub JPB4_GotFocus()
On Error Resume Next
Call c_blok(JPB4)
JPB4.Alignment = 0
End Sub

Private Sub JPB4_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789,.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
End If
End Sub

Private Sub JPB4_LostFocus()
On Error Resume Next
Call c_Kosong(JPB4)
JPB4.Alignment = 1

End Sub

Private Sub JPB5a_GotFocus()
On Error Resume Next
 Call c_blok(JPB5a)
 JPB5a.Alignment = 0
End Sub

Private Sub JPB5a_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789,.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If

End Sub

Private Sub JPB5a_LostFocus()
On Error Resume Next
Call c_Kosong(JPB5a)
JPB5a.Alignment = 1
End Sub

Private Sub JPB5b_GotFocus()
On Error Resume Next
Call c_blok(JPB5b)
End Sub

Private Sub JPB5b_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789,.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
End If

End Sub

Private Sub JPB5b_LostFocus()
On Error Resume Next
Call c_Kosong(JPB5b)
End Sub

Private Sub JPB7a_Change()
On Error Resume Next
JPB7d.Text = JPB7a.Text
End Sub

Private Sub JPB7a_GotFocus()
On Error Resume Next
Call c_blok(JPB7a)
End Sub

Private Sub JPB7a_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789,.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
End If

End Sub

Private Sub JPB7a_LostFocus()
On Error Resume Next
Call c_Kosong(JPB7a)
JPB7a.Alignment = 1
End Sub

Private Sub JPB7b_GotFocus()
On Error Resume Next
Call c_blok(JPB7b)
End Sub

Private Sub JPB7b_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789,.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
End If

End Sub

Private Sub JPB7b_LostFocus()
On Error Resume Next
Call c_Kosong(JPB7b)
JPB7b.Alignment = 1
End Sub

Private Sub JPB7c_GotFocus()
On Error Resume Next
Call c_blok(JPB7c)
End Sub

Private Sub JPB7c_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789,.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
End If

End Sub

Private Sub JPB7c_LostFocus()
On Error Resume Next
Call c_Kosong(JPB7c)
JPB7c.Alignment = 1
End Sub

Private Sub JPB7d_GotFocus()
On Error Resume Next
Call c_blok(JPB7d)
End Sub

Private Sub JPB7d_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789,.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If

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

Private Sub JPB7d_LostFocus()
On Error Resume Next
Call c_Kosong(JPB7d)
JPB7d.Alignment = 1
End Sub
