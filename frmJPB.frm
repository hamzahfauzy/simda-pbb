VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmJPB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Jenis Penggunaan Bangunan (JPB)"
   ClientHeight    =   8715
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12195
   ControlBox      =   0   'False
   Icon            =   "frmJPB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmJPB.frx":1CCA
   ScaleHeight     =   8715
   ScaleWidth      =   12195
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   90
      Picture         =   "frmJPB.frx":2994
      ScaleHeight     =   390
      ScaleWidth      =   11970
      TabIndex        =   31
      Top             =   -15
      Width           =   11970
      Begin VB.CheckBox chPajak 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   10275
         TabIndex        =   2
         Top             =   90
         Width           =   1665
      End
      Begin VB.CheckBox chPajak 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   6990
         TabIndex        =   0
         Top             =   90
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CheckBox chPajak 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   8505
         TabIndex        =   1
         Top             =   90
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JPB Bangunan Non Standard"
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
         Left            =   540
         TabIndex        =   32
         Top             =   60
         Width           =   2100
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   360
      Left            =   7890
      TabIndex        =   25
      Top             =   7620
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   750
      Left            =   7770
      TabIndex        =   21
      Top             =   270
      Width           =   4305
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   285
         Width           =   2520
      End
      Begin VB.CommandButton cmdCari 
         Height          =   360
         Left            =   3840
         Picture         =   "frmJPB.frx":6FFC
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "HARGA BARU"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   23
         Top             =   360
         Width           =   1365
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
      Left            =   6405
      TabIndex        =   12
      Top             =   8085
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
      Left            =   5430
      TabIndex        =   11
      Top             =   8085
      Width           =   990
   End
   Begin MSComctlLib.ListView vBangunan 
      Height          =   6465
      Left            =   75
      TabIndex        =   18
      Top             =   1035
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11404
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Kelas"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Lt. Min"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Lt. Max"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Nilai Lama"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Nilai Baru"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Ket"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
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
      Height          =   435
      Left            =   4455
      TabIndex        =   10
      Top             =   8085
      Width           =   990
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   90
      TabIndex        =   26
      Top             =   270
      Width           =   7680
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
         Left            =   3285
         TabIndex        =   4
         Top             =   285
         Width           =   4305
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
         Left            =   1350
         TabIndex        =   3
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label12 
         Caption         =   "JPB"
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
         Left            =   2835
         TabIndex        =   28
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label13 
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
         Height          =   225
         Left            =   135
         TabIndex        =   27
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   405
      TabIndex        =   15
      Top             =   8835
      Visible         =   0   'False
      Width           =   7740
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
         Height          =   300
         Index           =   3
         Left            =   1515
         TabIndex        =   19
         Top             =   525
         Width           =   6150
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
         Height          =   300
         Index           =   0
         Left            =   1515
         TabIndex        =   16
         Top             =   210
         Width           =   6150
      End
      Begin VB.Label Label1 
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
         Height          =   180
         Left            =   180
         TabIndex        =   20
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label41 
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
         Height          =   240
         Left            =   180
         TabIndex        =   17
         Top             =   240
         Width           =   1950
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   525
      Left            =   9510
      TabIndex        =   29
      Top             =   7410
      Width           =   2565
      Begin VB.CommandButton cmdPersen 
         BackColor       =   &H00C0FFFF&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   135
         Width           =   345
      End
      Begin VB.TextBox tPersen 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   1155
         TabIndex        =   6
         Text            =   "0"
         Top             =   165
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kenaikan (%)"
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
         TabIndex        =   30
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Height          =   525
      Left            =   90
      TabIndex        =   33
      Top             =   7410
      Width           =   9435
      Begin VB.CommandButton tTower 
         Height          =   345
         Left            =   8910
         Picture         =   "frmJPB.frx":7CC6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   105
         Width           =   405
      End
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   6615
         TabIndex        =   9
         Top             =   120
         Width           =   2235
      End
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3600
         TabIndex        =   8
         Top             =   120
         Width           =   1920
      End
      Begin VB.TextBox tBumi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Top             =   120
         Width           =   1905
      End
      Begin VB.Label Label7 
         Caption         =   "Mekanik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5700
         TabIndex        =   36
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "Pagar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2895
         TabIndex        =   35
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "Tower"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   75
         TabIndex        =   34
         Top             =   195
         Width           =   1365
      End
   End
   Begin MSDataGridLib.DataGrid gg 
      Height          =   3960
      Left            =   -480
      TabIndex        =   24
      Top             =   5280
      Visible         =   0   'False
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   6985
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   9180
      Left            =   -330
      Picture         =   "frmJPB.frx":8990
      Stretch         =   -1  'True
      Top             =   -420
      Width           =   12960
   End
End
Attribute VB_Name = "frmJPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jumRek, K1, K2, PBBMin
Dim xxTahun
Private Sub cmdBangunan_Click()
On Error Resume Next
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
On Error Resume Next
frmOP_Tanah.Show
End Sub


Private Sub cboNOP_GotFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 2
    callKomp1
End Select
End Sub

Private Sub cboNOP_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub cboNOP_LostFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case 1, 2
For i = 0 To cboNOP(Index).ListCount - 1
        If (UCase(cboNOP(Index).List(i)) Like "*" + UCase(cboNOP(Index).Text) + "*" = True) Then
            cboNOP(Index).Text = cboNOP(Index).List(i)
            cboNOP_Click (Index)
            Exit Sub
        End If
          If i = cboNOP(Index).ListCount - 1 Then
            If UCase(cboNOP(Index).List(i)) Like "*" + UCase(cboNOP(Index).Text) + "*" = False Then
                cboNOP(Index).Text = cboNOP(Index).List(0)
                cboNOP_Click (Index)
                Exit Sub
            End If
        End If
    Next
End Select
End Sub

Private Sub chPajak_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 1
    If chPajak(1).Value = 1 Then
        chPajak(2).Value = 0
        chPajak(3).Value = 0
        cmdSave.Caption = "&Save"
        xxTahun = (cboNOP(1).Text * 1) - 1
        bersih
    Else
        If chPajak(2).Value = 0 And chPajak(3).Value = 0 Then
            chPajak(1).Value = 1
        End If
    End If
Case 2
    
   If chPajak(2).Value = 1 Then
    chPajak(1).Value = 0
    chPajak(3).Value = 0
    cmdSave.Caption = "&Delete"
    xTanya = MsgBox("Apa anda yakin menghapus Nilai DBKB Non Standard?", vbQuestion + vbYesNo, "Penghapusan")
    
    If xTanya = vbYes Then
        cmdSave.Caption = "&Delete"
        xxTahun = cboNOP(1).Text * 1
        bersih
    Else
        chPajak(1).Value = 1
        chPajak(2).Value = 0
        cmdSave.Caption = "&Save"
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
        cmdSave.Caption = "&Update"
        xTanya = MsgBox("Apa anda yakin melakukan pemutakhiran Nilai DBKB Non Standard?", vbQuestion + vbYesNo, "Pemutakhiran")
        If xTanya = vbYes Then
            cmdSave.Caption = "&Update"
            xxTahun = cboNOP(1).Text * 1
            bersih
        Else
            chPajak(1).Value = 1
            chPajak(3).Value = 0
            cmdSave.Caption = "&Save"
        End If
    Else
        If chPajak(1).Value = 0 And chPajak(2).Value = 0 Then
            chPajak(3).Value = 1
        End If
    End If
End Select
End Sub

Private Sub chPajak_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
bersih
End Sub

Private Sub cmdCari_Click()
On Error GoTo Salah
vBangunan.SelectedItem.ListSubItems(6).Text = Format(tBumi(0).Text, "#,#0.00")
vBangunan.SelectedItem.ListSubItems(7).Text = "OK"
tBumi(0).Text = 0
vBangunan.SetFocus
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Description, vbCritical, "TETNONG: " & Err.Number

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdID_Click()
On Error Resume Next
xID = 1
frmList_Subjek.Show
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

'Private Sub cmdExport_Click()
'On Error Resume Next
'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Set oExcel = CreateObject("Excel.Application")
'Set oBook = oExcel.Workbooks.Add
'Set oSheet = oBook.Worksheets(1)
'On Error GoTo errcode
'With oBook.Worksheets("Sheet1").Rows(1)
'   .Font.Bold = True
'   For J = 0 To gg.Columns.Count - 1
'
'    Worksheets("sheet1").Cells(1, J + 1).Value = gg.Columns(J).Caption
'    Next
'End With
'oSheet.Range("A2").CopyFromRecordset rPajak
'oBook.SaveAs cboNOP(2).Text
'oBook.Close
'oExcel.Quit
'Exit Sub
'errcode:
'MsgBox Err.Number & " : " & Err.Source & "->" & Err.Description
'
'End Sub

Private Sub cmdPersen_Click()
On Error GoTo Salah
Dim JUMP
For i = 1 To vBangunan.ListItems.Count
    JUMP = (tPersen.Text * vBangunan.ListItems.Item(i).ListSubItems(5).Text) / 100
    If vBangunan.ListItems.Item(i).ListSubItems(7).Text = "-" Then
        vBangunan.ListItems.Item(i).ListSubItems(6).Text = Format(vBangunan.ListItems.Item(i).ListSubItems(5).Text + JUMP, "#,#0.00")
    End If
   'CALL_KELAS
Next
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cmdSave_Click()
On Error GoTo Salah
Dim Pesan, Judul
If cmdSave.Caption = "&Save" Then
    Pesan = "Apa anda yakin akan menyimpan data ? "
    Judul = "Saved..."
ElseIf cmdSave.Caption = "&Update" Then
    Pesan = "Data yang telah diubah akan disimpan (Update). Lanjutkan? "
    Judul = "Updated..."
Else
    Pesan = "Seluruh record yang tampil akan terhapus. Lanjutkan? "
    Judul = "Deleted..."
End If
TANYA = MsgBox(Pesan, vbInformation + vbYesNo, Judul)
If TANYA = vbYes Then
    If Left(Trim(cboNOP(2).Text), 2) = "02" Then
        SIMPAN_JPB2
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "03" Then
        SIMPAN_JPB3
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "04" Then
        SIMPAN_JPB4
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "05" Then
        SIMPAN_JPB5
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "06" Then
        SIMPAN_JPB6
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "07" Then
        SIMPAN_JPB7
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "08" Then
        SIMPAN_JPB8
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "09" Then
        SIMPAN_JPB9
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "10" Then
        'SIMPAN_JPB10
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "11" Then
        'SIMPAN_JPB11
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "12" Then
        SIMPAN_JPB12
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "13" Then
        SIMPAN_JPB13
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "14" Then
        SIMPAN_JPB14
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "15" Then
        SIMPAN_JPB15
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "16" Then
        SIMPAN_JPB16
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "17" Then
        SIMPAN_JPB17
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "18" Then
        SIMPAN_DAYA_DUKUNG
    ElseIf Left(Trim(cboNOP(2).Text), 2) = "19" Then
        SIMPAN_MEZANIN
    End If
Else
    Exit Sub
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim C(100)
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2

'For i = 0 To 4
'    cboNOP(i).Clear
'Next
'cbonop(2).Text = "0001"

callKec
'tBumi(1).Text = 0
'cboZNT.Clear
cboNOP(1).Clear
cboNOP(1).Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    cboNOP(1).AddItem i
Next
xxTahun = (cboNOP(1).Text * 1) - 1
Frame4.Visible = False
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
End Sub






Private Sub cboNOP_Click(Index As Integer)
'On Error Resume Next
On Error GoTo Salah
Select Case Index
    Case 0
            cboNOP(2).Clear
    Case 1
         cboNOP(2).Clear
         
    Case 2
        Frame4.Visible = False
        Frame1.Enabled = True
        Frame3.Enabled = True
        If Left(cboNOP(2).Text, 2) = 2 Then
            callJPB2
        ElseIf Left(cboNOP(2).Text, 2) = 3 Then
            callJPB3
        ElseIf Left(cboNOP(2).Text, 2) = 4 Then
            callJPB4
        ElseIf Left(cboNOP(2).Text, 2) = 5 Then
            callJPB5
        ElseIf Left(cboNOP(2).Text, 2) = 6 Then
            callJPB6
        ElseIf Left(cboNOP(2).Text, 2) = 7 Then
            callJPB7
        ElseIf Left(cboNOP(2).Text, 2) = 8 Then
            callJPB8
        ElseIf Left(cboNOP(2).Text, 2) = 9 Then
            callJPB9
        ElseIf Left(cboNOP(2).Text, 2) = 12 Then
            callJPB12
        ElseIf Left(cboNOP(2).Text, 2) = 13 Then
            callJPB13
        ElseIf Left(cboNOP(2).Text, 2) = 14 Then
            callJPB14
        ElseIf Left(cboNOP(2).Text, 2) = 15 Then
            callJPB15
        ElseIf Left(cboNOP(2).Text, 2) = 16 Then
            callJPB16
        ElseIf Left(cboNOP(2).Text, 2) = 17 Then
            Frame4.Visible = True
            Frame1.Enabled = False
            Frame3.Enabled = False
            callJPB17
        ElseIf Left(cboNOP(2).Text, 2) = 18 Then
            callJPB18
        ElseIf Left(cboNOP(2).Text, 2) = 19 Then
            callJPB19
        Else
            MsgBox "Bukan Bangunan Non Standar...", vbCritical, "Salah Pilih.."
            vBangunan.ListItems.Clear
            
        End If
    Case 3
        
        'tBumi(0).Text = K1 & "." & K2 & "." & Left(cboNOP(0).Text, 3) & "." & Left(cboNOP(1).Text, 3) & "." & Left(cbonop(1).Text, 3) & "-" & cbonop(2).Text & "." & Left(cboNOP(4).Text, 1)
End Select
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cboNOP_DropDown(Index As Integer)
On Error Resume Next
Select Case Index
    Case 1
    Case 2
    
    callKomp1
    Case 3
        callKel
        
End Select

End Sub





Sub callKec()
On Error GoTo Salah
cboNOP(0).Clear
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
Sub callKomp1()
On Error GoTo Salah
cboNOP(2).Clear ': cboZNT.Clear
STRITEM = "Select * From REF_JPB WHERE KD_JPB <> '01' AND KD_JPB <> '10' AND KD_JPB <> '11' order by KD_JPB asc"
        openDB (STRITEM)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
        cboNOP(2).AddItem rPajak!KD_JPB & "-" & rPajak!NM_JPB
rPajak.MoveNext
Loop
    cboNOP(2).AddItem "18-DBKB DAYA DUKUNG LANTAI"
    cboNOP(2).AddItem "19-DBKB MEZANIN"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub


Private Sub tBumi_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
End If
Select Case Index
Case 0
    If InStr("0123456789-.,", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
    If KeyAscii = 13 Then
        cmdCari_Click
        KeyAscii = 0
    End If
Case 1, 2, 3
    If InStr("0123456789-.,", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
End Select
End Sub

Private Sub tPersen_KeyPress(KeyAscii As Integer)
On Error Resume Next
If InStr("0123456789-.,", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
        If KeyAscii = 13 Then
            cmdPersen_Click
        End If
End Sub

Private Sub tTower_Click()
On Error GoTo Salah
Dim Jumlah
Jumlah = (tBumi(1).Text * 1) + (tBumi(2).Text * 1) + (tBumi(3).Text * 1)
vBangunan.SelectedItem.ListSubItems(3).Text = Format(tBumi(1).Text, "#,#0.00")
vBangunan.SelectedItem.ListSubItems(4).Text = Format(tBumi(2).Text, "#,#0.00")
vBangunan.SelectedItem.ListSubItems(5).Text = Format(tBumi(3).Text, "#,#0.00")
vBangunan.SelectedItem.ListSubItems(6).Text = Format(Jumlah, "#,#0.00")
vBangunan.SelectedItem.ListSubItems(7).Text = "OK"
tBumi(0).Text = 0
vBangunan.SetFocus
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Description, vbCritical, "TETNONG: " & Err.Number
End Sub

Private Sub vBangunan_Click()
On Error GoTo Salah
'If vBangunan.SelectedItem.ListSubItems(5).Text = 1 Then
    For i = 1 To vBangunan.ListItems.Count
        vBangunan.ListItems.Item(i).ListSubItems(8).Text = "-"
    Next
 vBangunan.SelectedItem.ListSubItems(8).Text = "Proses"
    tBumi(0).Text = vBangunan.SelectedItem.ListSubItems(5).Text
    tBumi(1).Text = vBangunan.SelectedItem.ListSubItems(3).Text
    tBumi(2).Text = vBangunan.SelectedItem.ListSubItems(4).Text
    tBumi(3).Text = vBangunan.SelectedItem.ListSubItems(5).Text
'    tBumi(1).Text = vBangunan.SelectedItem.ListSubItems(5).Text
    For i = 1 To vBangunan.ListItems.Count
            If vBangunan.ListItems.Item(i).ListSubItems(7).Text = "OK" Then
                vBangunan.ListItems.Item(i).ListSubItems(7).Text = "OK"
            Else
                vBangunan.ListItems.Item(i).ListSubItems(7).Text = "-"
            End If
            
    Next
 
            If vBangunan.SelectedItem.ListSubItems(7).Text = "OK" Then
                Exit Sub
                vBangunan.SetFocus
            Else
               ' tBumi(0).SetFocus
            End If
    'End If
'    If vBangunan.SelectedItem.ListSubItems(5).Text * 1 = vBangunan.SelectedItem.ListSubItems(6).Text * 1 Then
'        vBangunan.SelectedItem.ListSubItems(7).Text = "Proses"
'        tBumi(0).SetFocus
'    Else
'        vBangunan.SelectedItem.ListSubItems(7).Text = "Sudah Proses"
'    End If
    
'End If
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Description, vbCritical, "TETNONG: " & Err.Number
End Sub

Private Sub vBangunan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan.SortKey = ColumnHeader.Index - 1
vBangunan.Sorted = True
vBangunan.Sorted = False
vBangunan.SortOrder = lvwAscending
End Sub

Sub callKel()
On Error GoTo Salah
cboNOP(3).Clear ': cboZNT.Clear
strKEL = "Select * From REF_KELURAHAN where KD_KECAMATAN='" & Left(cboNOP(0).Text, 3) & "' order by KD_KELURAHAN"
openDB (strKEL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
cboNOP(3).AddItem rPajak!KD_KELURAHAN & "-" & rPajak!NM_KELURAHAN
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Sub callJPB2()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB2 where  THN_DBKB_JPB2 = '" & xxTahun & "' order by KLS_DBKB_JPB2 asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!KLS_DBKB_JPB2)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak!LANTAI_MIN_JPB2)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!LANTAI_MAX_JPB2)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB2), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(3).Text = "Kelas"
        vBangunan.ColumnHeaders(4).Text = "LT. Min"
        vBangunan.ColumnHeaders(5).Text = "LT. Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 0
        vBangunan.ColumnHeaders(3).Width = 1000
        vBangunan.ColumnHeaders(4).Width = 1000
        vBangunan.ColumnHeaders(5).Width = 1000
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
Sub callJPB3()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB3 where THN_DBKB_JPB3 = '" & xxTahun & "' order by TING_KOLOM_MAX_DBKB_JPB3,TING_KOLOM_MIN_DBKB_JPB3,LBR_BENT_MAX_DBKB_JPB3,LBR_BENT_MIN_DBKB_JPB3 asc"
        openDB (STRITEM)

        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Trim(rPajak!LBR_BENT_MIN_DBKB_JPB3)
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!LBR_BENT_MAX_DBKB_JPB3)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak!TING_KOLOM_MIN_DBKB_JPB3)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!TING_KOLOM_MAX_DBKB_JPB3)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB3), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(2).Text = "LB Min"
        vBangunan.ColumnHeaders(3).Text = "LB Max"
        vBangunan.ColumnHeaders(4).Text = "TK Min"
        vBangunan.ColumnHeaders(5).Text = "TK Max"
        vBangunan.ColumnHeaders(2).Width = 1000
        vBangunan.ColumnHeaders(3).Width = 1000
        vBangunan.ColumnHeaders(4).Width = 1000
        vBangunan.ColumnHeaders(5).Width = 1000
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
Set gg.DataSource = rPajak
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub
Sub callJPB4()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB4 where  THN_DBKB_JPB4 = '" & xxTahun & "' order by KLS_DBKB_JPB4 asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!KLS_DBKB_JPB4)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak!LANTAI_MIN_JPB4)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!LANTAI_MAX_JPB4)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB4), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(3).Text = "Kelas"
        vBangunan.ColumnHeaders(4).Text = "LT. Min"
        vBangunan.ColumnHeaders(5).Text = "LT. Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 0
        vBangunan.ColumnHeaders(3).Width = 1000
        vBangunan.ColumnHeaders(4).Width = 1000
        vBangunan.ColumnHeaders(5).Width = 1000
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub
Sub callJPB5()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB5 where  THN_DBKB_JPB5 = '" & xxTahun & "' order by KLS_DBKB_JPB5 asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!KLS_DBKB_JPB5)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak!LANTAI_MIN_JPB5)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!LANTAI_MAX_JPB5)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB5), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(3).Text = "Kelas"
        vBangunan.ColumnHeaders(4).Text = "LT. Min"
        vBangunan.ColumnHeaders(5).Text = "LT. Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 0
        vBangunan.ColumnHeaders(3).Width = 1000
        vBangunan.ColumnHeaders(4).Width = 1000
        vBangunan.ColumnHeaders(5).Width = 1000
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub


Sub callJPB6()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB6 where  THN_DBKB_JPB6 = '" & xxTahun & "' order by KLS_DBKB_JPB6 asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", 0 'Trim(rPajak!KLS_DBKB_JPB5)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", 0 'Trim(rPajak!LANTAI_MIN_JPB5)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!KLS_DBKB_JPB6)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB6), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(2).Text = "Kelas"
        vBangunan.ColumnHeaders(3).Text = "Kelas"
        vBangunan.ColumnHeaders(4).Text = "LT. Min"
        vBangunan.ColumnHeaders(5).Text = "LT. Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 1000
        vBangunan.ColumnHeaders(3).Width = 0
        vBangunan.ColumnHeaders(4).Width = 0
        vBangunan.ColumnHeaders(5).Width = 0
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
    
End Sub
Sub callJPB7()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB7 where THN_DBKB_JPB7 = '" & xxTahun & "' order by JNS_DBKB_JPB7 asc"
        openDB (STRITEM)

        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Trim(rPajak!JNS_DBKB_JPB7)
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!BINTANG_DBKB_JPB7)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak!LANTAI_MIN_JPB7)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!LANTAI_MAX_JPB7)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB7), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(2).Text = "Jenis"
        vBangunan.ColumnHeaders(3).Text = "Bintang"
        vBangunan.ColumnHeaders(4).Text = "LT Min"
        vBangunan.ColumnHeaders(5).Text = "LT Max"
        vBangunan.ColumnHeaders(2).Width = 1000
        vBangunan.ColumnHeaders(3).Width = 1000
        vBangunan.ColumnHeaders(4).Width = 1000
        vBangunan.ColumnHeaders(5).Width = 1000
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub

Sub callJPB8()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB8 where THN_DBKB_JPB8 = '" & xxTahun & "' order by TING_KOLOM_MAX_DBKB_JPB8,TING_KOLOM_MIN_DBKB_JPB8,LBR_BENT_MAX_DBKB_JPB8,LBR_BENT_MIN_DBKB_JPB8 asc"
        openDB (STRITEM)

        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Trim(rPajak!LBR_BENT_MIN_DBKB_JPB8)
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!LBR_BENT_MAX_DBKB_JPB8)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak!TING_KOLOM_MIN_DBKB_JPB8)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!TING_KOLOM_MAX_DBKB_JPB8)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB8), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(2).Text = "LB Min"
        vBangunan.ColumnHeaders(3).Text = "LB Max"
        vBangunan.ColumnHeaders(4).Text = "TK Min"
        vBangunan.ColumnHeaders(5).Text = "TK Max"
        vBangunan.ColumnHeaders(2).Width = 1000
        vBangunan.ColumnHeaders(3).Width = 1000
        vBangunan.ColumnHeaders(4).Width = 1000
        vBangunan.ColumnHeaders(5).Width = 1000
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub
Sub callJPB9()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB9 where  THN_DBKB_JPB9 = '" & xxTahun & "' order by KLS_DBKB_JPB9 asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!KLS_DBKB_JPB9)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak!LANTAI_MIN_JPB9)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!LANTAI_MAX_JPB9)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB9), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(3).Text = "Kelas"
        vBangunan.ColumnHeaders(4).Text = "LT. Min"
        vBangunan.ColumnHeaders(5).Text = "LT. Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 0
        vBangunan.ColumnHeaders(3).Width = 1000
        vBangunan.ColumnHeaders(4).Width = 1000
        vBangunan.ColumnHeaders(5).Width = 1000
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub
Sub callJPB12()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB12 where  THN_DBKB_JPB12 = '" & xxTahun & "' order by TYPE_DBKB_JPB12 asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", 0 'Trim(rPajak!KLS_DBKB_JPB5)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", 0 'Trim(rPajak!LANTAI_MIN_JPB5)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!TYPE_DBKB_JPB12)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB12), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(2).Text = "TYPE"
        vBangunan.ColumnHeaders(3).Text = "Kelas"
        vBangunan.ColumnHeaders(4).Text = "LT. Min"
        vBangunan.ColumnHeaders(5).Text = "LT. Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 1000
        vBangunan.ColumnHeaders(3).Width = 0
        vBangunan.ColumnHeaders(4).Width = 0
        vBangunan.ColumnHeaders(5).Width = 0
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub

Sub callJPB13()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB13 where  THN_DBKB_JPB13 = '" & xxTahun & "' order by KLS_DBKB_JPB13 asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!KLS_DBKB_JPB13)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak!LANTAI_MIN_JPB13)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!LANTAI_MAX_JPB13)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB13), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(3).Text = "Kelas"
        vBangunan.ColumnHeaders(4).Text = "LT. Min"
        vBangunan.ColumnHeaders(5).Text = "LT. Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 0
        vBangunan.ColumnHeaders(3).Width = 1000
        vBangunan.ColumnHeaders(4).Width = 1000
        vBangunan.ColumnHeaders(5).Width = 1000
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub

Sub callJPB14()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB14 where  THN_DBKB_JPB14 = '" & xxTahun & "' order by THN_DBKB_JPB14 asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", 0 'Trim(rPajak!KLS_DBKB_JPB13)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", 0 'Trim(rPajak!LANTAI_MIN_JPB13)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", 0 'Trim(rPajak!LANTAI_MAX_JPB13)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB14), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(3).Text = "Kelas"
        vBangunan.ColumnHeaders(4).Text = "LT. Min"
        vBangunan.ColumnHeaders(5).Text = "LT. Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 0
        vBangunan.ColumnHeaders(3).Width = 0
        vBangunan.ColumnHeaders(4).Width = 0
        vBangunan.ColumnHeaders(5).Width = 0
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub



Sub callJPB15()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB15 where  THN_DBKB_JPB15 = '" & xxTahun & "' order by JNS_TANGKI_DBKB_JPB15 asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!JNS_TANGKI_DBKB_JPB15)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak!KAPASITAS_MIN_DBKB_JPB15)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!KAPASITAS_MAX_DBKB_JPB15)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB15), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(3).Text = "Jenis Tangki"
        vBangunan.ColumnHeaders(4).Text = "Kap. Min"
        vBangunan.ColumnHeaders(5).Text = "Kap Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 1000
        vBangunan.ColumnHeaders(3).Width = 1000
        vBangunan.ColumnHeaders(4).Width = 1000
        vBangunan.ColumnHeaders(5).Width = 1000
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub
Sub callJPB16()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB16 where  THN_DBKB_JPB16 = '" & xxTahun & "' order by KLS_DBKB_JPB16 asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!KLS_DBKB_JPB16)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak!LANTAI_MIN_JPB16)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak!LANTAI_MAX_JPB16)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_JPB16), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(3).Text = "Kelas"
        vBangunan.ColumnHeaders(4).Text = "LT. Min"
        vBangunan.ColumnHeaders(5).Text = "LT. Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 0
        vBangunan.ColumnHeaders(3).Width = 1000
        vBangunan.ColumnHeaders(4).Width = 1000
        vBangunan.ColumnHeaders(5).Width = 1000
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub
Sub callJPB17()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_JPB17 where  THN_DBKB_JPB17= '" & xxTahun & "' "
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Trim(rPajak!TINGGI_MIN_JPB17) '"MENARA/TOWER"
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!TINGGI_MAX_JPB17) '"MENARA/TOWER"
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Format(Trim(rPajak!NILAI_BNG_MENARA_JPB17), "#,#0.00") 'Trim(rPajak!TINGGI_MIN_JPB17) '& " s.d " & Trim(rPajak!TINGGI_MAX_JPB17)
        vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Format(Trim(rPajak!NILAI_BGN_PAGAR_JPB17), "#,#0.00") 'Trim(rPajak!TINGGI_MAX_JPB17)
        vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!BIAYA_MEKANIK_JPB17), "#,#0.00")
        vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", Format(Trim(rPajak!NILAI_DBKB_JPB17), "#,#0.00")
        vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", 0
        vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
'        xLAST1 = I
'        rPajak.MoveNext
'        Loop
'        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'        J = xLAST1
'        Do While Not rPajak.EOF
'        J = J + 1
'        vBangunan.ListItems.Add J, "", Format(J, "#0")
'        vBangunan.ListItems.Item(J).ListSubItems.Add 1, "", Format(J, "#0")
'        vBangunan.ListItems.Item(J).ListSubItems.Add 2, "", "PAGAR"
'        vBangunan.ListItems.Item(J).ListSubItems.Add 3, "", Trim(rPajak!TINGGI_MIN_JPB17) '& " s.d " & Trim(rPajak!TINGGI_MAX_JPB17)
'        vBangunan.ListItems.Item(J).ListSubItems.Add 4, "", Trim(rPajak!TINGGI_MAX_JPB17) 'rPajak!BIAYA_MEKANIK_JPB17
'        vBangunan.ListItems.Item(J).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_BGN_PAGAR_JPB17), "#,#0.00")
'        vBangunan.ListItems.Item(J).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
'        vBangunan.ListItems.Item(J).ListSubItems.Add 7, "", "-"
'        vBangunan.ListItems.Item(J).ListSubItems.Add 8, "", "-"
'        xLAST2 = J
'        'End If
'        rPajak.MoveNext
'        Loop
'
'        If rPajak.RecordCount > 0 Then rPajak.MoveLast
'        K = xLAST2 + 1
'        Do While Not rPajak.EOF
'
'        vBangunan.ListItems.Add K, "", Format(K, "#0")
'        vBangunan.ListItems.Item(K).ListSubItems.Add 1, "", Format(K, "#0")
'        vBangunan.ListItems.Item(K).ListSubItems.Add 2, "", "BIAYA MECHANICAL ELECTRICAL"
'        vBangunan.ListItems.Item(K).ListSubItems.Add 3, "", "-"
'        vBangunan.ListItems.Item(K).ListSubItems.Add 4, "", "-" 'rPajak!BIAYA_MEKANIK_JPB17
'        vBangunan.ListItems.Item(K).ListSubItems.Add 5, "", Format(Trim(rPajak!BIAYA_MEKANIK_JPB17), "#,#0.00")
'        vBangunan.ListItems.Item(K).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
'        vBangunan.ListItems.Item(K).ListSubItems.Add 7, "", "-"
'        vBangunan.ListItems.Item(K).ListSubItems.Add 8, "", "-"
'
'        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(2).Text = "T_Min(M)"
        vBangunan.ColumnHeaders(3).Text = "T_Max(M)"
        vBangunan.ColumnHeaders(4).Text = "BNG_MENARA"
        vBangunan.ColumnHeaders(5).Text = "BNG_PAGAR"
        vBangunan.ColumnHeaders(6).Text = "Biaya Mekanik"
        vBangunan.ColumnHeaders(7).Text = "DBKB (JUMLAH)"
        'vBangunan.ColumnHeaders(8).Text = "DBKB Baru"
        vBangunan.ColumnHeaders(1).Width = 0
        vBangunan.ColumnHeaders(2).Width = 1100
        vBangunan.ColumnHeaders(3).Width = 1100
        vBangunan.ColumnHeaders(4).Width = 1600
        vBangunan.ColumnHeaders(5).Width = 1600
        vBangunan.ColumnHeaders(6).Width = 1600
        vBangunan.ColumnHeaders(7).Width = 1600
        'vBangunan.ColumnHeaders(8).Width = 1600
'        vBangunan.ColumnHeaders(1).Alignment = lvwColumnCenter
        vBangunan.ColumnHeaders(2).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub
Sub callJPB18()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_DAYA_DUKUNG where  THN_DBKB_DAYA_DUKUNG = '" & xxTahun & "' order by TYPE_KONSTRUKSI asc"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Trim(rPajak!TYPE_KONSTRUKSI)
        If Trim(rPajak!TYPE_KONSTRUKSI) = 1 Then
                    KX = "RINGAN"
                    DDMin = 1
                    DDMax = 600
                ElseIf Trim(rPajak!TYPE_KONSTRUKSI) = 2 Then
                    KX = "SEDANG"
                    DDMin = 601
                    DDMax = 1200
                ElseIf Trim(rPajak!TYPE_KONSTRUKSI) = 3 Then
                    KX = "MENENGAH"
                    DDMin = 1201
                    DDMax = 2400
                ElseIf Trim(rPajak!TYPE_KONSTRUKSI) = 4 Then
                    KX = "BERAT"
                    DDMin = 2401
                    DDMax = 5000
                Else
                    KX = "SANGAT BERAT"
                    DDMin = 5001
                    DDMax = 9999
                End If
                
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", KX 'Trim(rPajak!TYPE_KONSTRUKSI)
                
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", DDMin 'Trim(rPajak!LANTAI_MIN_JPB16)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", DDMax
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_DAYA_DUKUNG), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(2).Text = "Tipe"
        vBangunan.ColumnHeaders(3).Text = "Konstruksi"
        vBangunan.ColumnHeaders(4).Text = "Daya Min"
        vBangunan.ColumnHeaders(5).Text = "Daya Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 1000
        vBangunan.ColumnHeaders(3).Width = 1500
        vBangunan.ColumnHeaders(4).Width = 1600
        vBangunan.ColumnHeaders(5).Width = 1600
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnLeft
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub



Sub callJPB19()
On Error GoTo Salah
vBangunan.ListItems.Clear
        STRITEM = "Select * From DBKB_MEZANIN where  THN_DBKB_MEZANIN = '" & xxTahun & "'"
        openDB (STRITEM)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#0")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", 0 'Trim(rPajak!KLS_DBKB_JPB5)
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", 0 'Trim(rPajak!LANTAI_MIN_JPB5)
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", 0 'Trim(rPajak!KLS_DBKB_JPB6)
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak!NILAI_DBKB_MEZANIN), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0 'Format(Trim(rPajak![HRG_RESOURCE]) , "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
        vBangunan.ColumnHeaders(2).Text = "No"
        vBangunan.ColumnHeaders(3).Text = "Kelas"
        vBangunan.ColumnHeaders(4).Text = "LT. Min"
        vBangunan.ColumnHeaders(5).Text = "LT. Max"
        vBangunan.ColumnHeaders(6).Text = "Nilai Lama"
        vBangunan.ColumnHeaders(7).Text = "Nilai Baru"
        vBangunan.ColumnHeaders(2).Width = 1000
        vBangunan.ColumnHeaders(3).Width = 0
        vBangunan.ColumnHeaders(4).Width = 0
        vBangunan.ColumnHeaders(5).Width = 0
        vBangunan.ColumnHeaders(3).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(4).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(5).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(6).Alignment = lvwColumnRight
        vBangunan.ColumnHeaders(7).Alignment = lvwColumnRight
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
End Sub

Private Sub vBangunan_KeyDown(KeyCode As Integer, Shift As Integer)
vBangunan_Click
End Sub

Private Sub vBangunan_KeyUp(KeyCode As Integer, Shift As Integer)
vBangunan_Click
End Sub
Sub bersih()
On Error Resume Next
tBumi(0).Text = 0
tBumi(1).Text = 0
tBumi(2).Text = 0
tBumi(3).Text = 0

vBangunan.ListItems.Clear
End Sub
Sub SIMPAN_JPB2()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB2 WHERE THN_DBKB_JPB2 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) Non Standard : " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB2 = cboNOP(1).Text
        rPajak!KLS_DBKB_JPB2 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!LANTAI_MIN_JPB2 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!LANTAI_MAX_JPB2 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB2 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub
Sub SIMPAN_JPB3()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB3 WHERE THN_DBKB_JPB3 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB3/Tangki : " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB3 = cboNOP(1).Text
        rPajak!LBR_BENT_MIN_DBKB_JPB3 = vBangunan.ListItems.Item(i).ListSubItems(1).Text
        rPajak!LBR_BENT_MAX_DBKB_JPB3 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!TING_KOLOM_MIN_DBKB_JPB3 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!TING_KOLOM_MAX_DBKB_JPB3 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB3 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub



Sub SIMPAN_JPB4()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB4 WHERE THN_DBKB_JPB4 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB4/Ruko,Apotik,..: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB4 = cboNOP(1).Text
        rPajak!KLS_DBKB_JPB4 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!LANTAI_MIN_JPB4 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!LANTAI_MAX_JPB4 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB4 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Sub SIMPAN_JPB5()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB5 WHERE THN_DBKB_JPB5 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB5/Rumah Sakit,..: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB5 = cboNOP(1).Text
        rPajak!KLS_DBKB_JPB5 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!LANTAI_MIN_JPB5 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!LANTAI_MAX_JPB5 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB5 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub


Sub SIMPAN_JPB6()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB6 WHERE THN_DBKB_JPB6 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB6/Olahraga,..: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB6 = cboNOP(1).Text
        rPajak!KLS_DBKB_JPB6 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB6 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Sub SIMPAN_JPB7()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB7 WHERE THN_DBKB_JPB7 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB7 (Hotel/Wisma) : " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
       rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB7 = cboNOP(1).Text
        rPajak!JNS_DBKB_JPB7 = vBangunan.ListItems.Item(i).ListSubItems(1).Text
        rPajak!BINTANG_DBKB_JPB7 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!LANTAI_MIN_JPB7 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!LANTAI_MAX_JPB7 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB7 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Sub SIMPAN_JPB8()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB8 WHERE THN_DBKB_JPB8 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB8/Bengkel: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB8 = cboNOP(1).Text
        rPajak!LBR_BENT_MIN_DBKB_JPB8 = vBangunan.ListItems.Item(i).ListSubItems(1).Text
        rPajak!LBR_BENT_MAX_DBKB_JPB8 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!TING_KOLOM_MIN_DBKB_JPB8 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!TING_KOLOM_MAX_DBKB_JPB8 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB8 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub


Sub SIMPAN_JPB9()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB9 WHERE THN_DBKB_JPB9 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB12 (Parkir)..: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB9 = cboNOP(1).Text
        rPajak!KLS_DBKB_JPB9 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!LANTAI_MIN_JPB9 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!LANTAI_MAX_JPB9 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB9 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub


Sub SIMPAN_JPB12()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB12 WHERE THN_DBKB_JPB12 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB9 (Gedung Pemerintah)..: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB12 = cboNOP(1).Text
        rPajak!TYPE_DBKB_JPB12 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB12 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Sub SIMPAN_JPB13()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB13 WHERE THN_DBKB_JPB13 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB13(Apartemen)..: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB13 = cboNOP(1).Text
        rPajak!KLS_DBKB_JPB13 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!LANTAI_MIN_JPB13 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!LANTAI_MAX_JPB13 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB13 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub


Sub SIMPAN_JPB14()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB14 WHERE THN_DBKB_JPB14 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB14(Pompa Bensin)..: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB14 = cboNOP(1).Text
        rPajak!NILAI_DBKB_JPB14 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Sub SIMPAN_JPB15()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB15 WHERE THN_DBKB_JPB15 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB15(Tangki Minyak)..: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB15 = cboNOP(1).Text
        rPajak!JNS_TANGKI_DBKB_JPB15 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!KAPASITAS_MIN_DBKB_JPB15 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!KAPASITAS_MAX_DBKB_JPB15 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB15 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub




Sub SIMPAN_JPB16()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB16 WHERE THN_DBKB_JPB16 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB16 (Gedung Sekolah)..: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB16 = cboNOP(1).Text
        rPajak!KLS_DBKB_JPB16 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!LANTAI_MIN_JPB16 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!LANTAI_MAX_JPB16 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB16 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Sub SIMPAN_JPB17()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_JPB17 WHERE THN_DBKB_JPB17 ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) JPB17 (TOWER)..: " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_JPB17 = cboNOP(1).Text
        rPajak!TINGGI_MIN_JPB17 = vBangunan.ListItems.Item(i).ListSubItems(1).Text
        rPajak!TINGGI_MAX_JPB17 = vBangunan.ListItems.Item(i).ListSubItems(2).Text
        rPajak!NILAI_BNG_MENARA_JPB17 = vBangunan.ListItems.Item(i).ListSubItems(3).Text
        rPajak!BIAYA_MEKANIK_JPB17 = vBangunan.ListItems.Item(i).ListSubItems(5).Text
        rPajak!NILAI_BGN_PAGAR_JPB17 = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak!NILAI_DBKB_JPB17 = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub



Sub SIMPAN_DAYA_DUKUNG()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_DAYA_DUKUNG WHERE THN_DBKB_DAYA_DUKUNG  ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) Daya Dukung Lantai Tahun : " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_DAYA_DUKUNG = cboNOP(1).Text
        rPajak!TYPE_KONSTRUKSI = vBangunan.ListItems.Item(i).ListSubItems(1).Text
        rPajak!NILAI_DBKB_DAYA_DUKUNG = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub


Sub SIMPAN_MEZANIN()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_MEZANIN where THN_DBKB_MEZANIN ='" & Trim(cboNOP(1).Text) & "' "
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) MEZANIN Tahun " & cboNOP(1).Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If (vBangunan.ListItems.Item(J).ListSubItems(6).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(6).Text = 0) And (vBangunan.ListItems.Item(J).ListSubItems(7).Text <> "OK") Then
        MsgBox "Nilai DBKB tidak lengkap, Silahkan dilengkapi terlebih dahulu...", vbCritical, "Tetnong"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" And chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!THN_DBKB_MEZANIN = cboNOP(1).Text
        rPajak!NILAI_DBKB_MEZANIN = vBangunan.ListItems.Item(i).ListSubItems(6).Text

    rPajak.Update
   
    
Next
Keluar:
vBangunan.ListItems.Clear
'chPajak(1).Value = 1
'cmdSave.Caption = "&Save"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

