VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNJOPTKP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penentuan NJOPTKP"
   ClientHeight    =   2385
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   4125
   ControlBox      =   0   'False
   Icon            =   "frmNJOPTKP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4125
   Begin MSComctlLib.ProgressBar pNilai 
      Height          =   255
      Left            =   90
      TabIndex        =   20
      Top             =   405
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -15
      ScaleHeight     =   360
      ScaleWidth      =   6225
      TabIndex        =   19
      Top             =   2055
      Width           =   6225
   End
   Begin VB.CheckBox cShow 
      Appearance      =   0  'Flat
      Caption         =   "Klik Untuk Menampilkan Data"
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
      Height          =   225
      Left            =   -15
      TabIndex        =   14
      Top             =   1860
      Visible         =   0   'False
      Width           =   3825
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
      Left            =   1935
      TabIndex        =   7
      Top             =   1560
      Width           =   930
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
      Left            =   2790
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
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
      Left            =   1035
      TabIndex        =   5
      Top             =   1560
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
      Height          =   825
      Left            =   75
      TabIndex        =   8
      Top             =   615
      Width           =   3975
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
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1650
         Width           =   2370
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   1320
         Width           =   1440
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
         Height          =   255
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         Top             =   1050
         Width           =   2355
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
         Left            =   1890
         TabIndex        =   0
         Top             =   300
         Width           =   1350
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   15
         X2              =   5670
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         X1              =   30
         X2              =   5685
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   5655
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   45
         X2              =   5700
         Y1              =   915
         Y2              =   915
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
         TabIndex        =   12
         Top             =   1710
         Width           =   1305
      End
      Begin VB.Label LNilai 
         BackColor       =   &H8000000B&
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
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   75
         TabIndex        =   4
         Top             =   2370
         Width           =   3870
      End
      Begin VB.Label Label3 
         Caption         =   "Objek Ke"
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
         Left            =   1560
         TabIndex        =   11
         Top             =   1350
         Width           =   1305
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
         Height          =   210
         Left            =   645
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Jumlah Objek"
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
         TabIndex        =   9
         Top             =   1035
         Width           =   1305
      End
   End
   Begin MSComctlLib.ListView vBangunan 
      Height          =   3285
      Left            =   4185
      TabIndex        =   13
      Top             =   315
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   5794
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
      Appearance      =   0
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
      NumItems        =   12
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
         Text            =   "ID SUBJEK"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "PROP"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "KAB"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "KEC"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "KEL"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "BLOK"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "URUT"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "JNS"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "TAHUN"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "JUMLAH"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView vBangunan1 
      Height          =   4020
      Left            =   60
      TabIndex        =   15
      Top             =   3945
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   7091
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
      Appearance      =   0
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
      NumItems        =   12
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
         Text            =   "ID SUBJEK"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "PROP"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "KAB"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "KEC"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "KEL"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "BLOK"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "URUT"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "JNS"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "TAHUN"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "JUMLAH"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   6225
      TabIndex        =   18
      Top             =   0
      Width           =   6225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "HASIL PEMILIHAN OBJEK PAJAK YANG MENDAPAT NJOPTKP"
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
      Left            =   6240
      TabIndex        =   17
      Top             =   60
      Width           =   4320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "DAFTAR SELURUH OBJEK PAJAK"
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
      TabIndex        =   16
      Top             =   3675
      Width           =   2340
   End
End
Attribute VB_Name = "frmNJOPTKP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Titik(4)
Dim Data1(100000), Data2(100000), Data3(100000), Data4(100000), Data5(100000), Data6(100000), Data7(100000), Data8(100000)

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
tJum.Text = 0
xPro(0).Text = 0
xPro(1).Text = ""
LNilai.Visible = False
vBangunan.ListItems.Clear
cShow.Value = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Salah
Screen.MousePointer = vbHourglass

Select Case ccMenu
Case 1
    cPesan = MsgBox("Posting SPPT tahun " & ccTahun.Text, vbQuestion + vbYesNo, "Posting...")
    If cPesan = vbYes Then
        dl_str = "Delete From SPPT_1 where THN_PAJAK_SPPT='" & ccTahun.Text & "'"
        openDB (dl_str)
        bc_STR = "INSERT INTO SPPT_1 SELECT * FROM SPPT WHERE SPPT.THN_PAJAK_SPPT='" & ccTahun.Text & "'"
        openDB (bc_STR)
        MsgBox "Sukses...!"
    End If

Case 2
n_STR = "select * FROM DAT_SUBJEK_PAJAK_NJOPTKP WHERE THN_NJOPTKP='" & ccTahun.Text & "' ORDER BY SUBJEK_PAJAK_ID ASC"
openDB (n_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then
    TANYA = MsgBox("NJOPTKP Tahun " & ccTahun.Text & " sudah ada" & _
            vbCrLf & "Ingin diulang kembali?", vbCritical + vbYesNo, "Tetnong...")
    If TANYA = vbNo Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End If

TANYA = MsgBox("NJOPTKP lama akah dihapus" & _
            vbCrLf & "Setuju...?", vbInformation + vbYesNo, "Processed...")
If TANYA = vbYes Then
   C_NJOPTKP
    pNilai.Visible = True
    For i = 1 To 20
            pNilai.Value = i
    Next
    'C_STR = "xx_NJOPTKP '" & ccTahun.Text & "'"
    'openDB (C_STR)
    For i = 21 To 100
            pNilai.Value = i
    Next
   sv_NJOPTKP
    MsgBox "Proses Sukses!"
    pNilai.Visible = False
    'dbPajak.Close
    'rPajak.Clse
End If
End Select
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub cShow_Click()
On Error Resume Next
If cShow.Value = 1 Then
    Me.Width = 12630
    Me.Height = 8445
Else
    Me.Width = 4254
    Me.Height = 4065
End If
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2

End Sub

Private Sub Form_Activate()
On Error Resume Next
Screen.MousePointer = vbHourglass
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
ccTahun.Clear
ccTahun.Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    ccTahun.AddItem i
Next
LNilai.Visible = False
pNilai.Visible = False
If ccMenu = 1 Then
    Me.Caption = "Pemindahan SPPT Lama"
Else
    Me.Caption = "Penentuan NJOPTKP"
End If
Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
xxID = 0
End Sub
Private Sub vBangunan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan.SortKey = ColumnHeader.Index - 1
vBangunan.Sorted = True
vBangunan.Sorted = False
vBangunan.SortOrder = lvwAscending
End Sub


Sub C_NJOPTKP()
On Error GoTo Salah
Dim cRec
Screen.MousePointer = vbHourglass
vBangunan.ListItems.Clear
vBangunan1.ListItems.Clear
'Q_STR = "SELECT QOBJEKPAJAK.SUBJEK_PAJAK_ID, QOBJEKPAJAK.JNS_BUMI From QOBJEKPAJAK GROUP BY QOBJEKPAJAK.SUBJEK_PAJAK_ID, QOBJEKPAJAK.JNS_BUMI HAVING (((QOBJEKPAJAK.JNS_BUMI)='1'))"
Q_STR = "SELECT SUBJEK_PAJAK_ID, JNS_BUMI From QOBJEKPAJAK GROUP BY SUBJEK_PAJAK_ID, JNS_BUMI HAVING JNS_BUMI='1'"
openDB (Q_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
cRec = 0
Do While Not rPajak.EOF
cRec = cRec + 1
rPajak.MoveNext
Loop

'n_str = "select subjek_pajak_id,NOPQ,NJOP_BUMI,NJOP_BNG from QOBJEKPAJAK WHERE JNS_BUMI='1' ORDER BY NJOP_BUMI+NJOP_BNG,SUBJEK_PAJAK_ID ASC"
'n_STR = "select subjek_pajak_id,jns_bumi from QOBJEKPAJAK WHERE JNS_BUMI='1' Group by SUBJEK_PAJAK_ID,JNS_BUMI ORDER BY SUBJEK_PAJAK_ID ASC"
'Q_STR = "SELECT QOBJEKPAJAK.SUBJEK_PAJAK_ID, QOBJEKPAJAK.JNS_BUMI From QOBJEKPAJAK GROUP BY QOBJEKPAJAK.SUBJEK_PAJAK_ID, QOBJEKPAJAK.JNS_BUMI HAVING (((QOBJEKPAJAK.JNS_BUMI)='1'))"
Q_STR = "SELECT SUBJEK_PAJAK_ID, JNS_BUMI From QOBJEKPAJAK GROUP BY SUBJEK_PAJAK_ID, JNS_BUMI HAVING JNS_BUMI='1'"

'n_STR = "select subjek_pajak_id,jns_bumi,KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP  from QOBJEKPAJAK WHERE JNS_BUMI='1' GROUP BY SUBJEK_PAJAK_ID,JNS_BUMI,KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP  ORDER BY SUBJEK_PAJAK_ID ASC "
openDB (Q_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'CQ = 0
'Do While Not rPajak.EOF
'    CQ = CQ + 1
'rPajak.MoveNext
'Loop
'If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0: J = 0
tJum.Text = cRec
'tJum.Text = 2995 'Pajak.RecordCount
'If rPajak.RecordCount > 1 Then
    pNilai.Max = tJum.Text
    pNilai.Min = 1
'End If
    Do While Not rPajak.EOF
        i = i + 1
        'J = J + 1
        LNilai.Visible = True
        LNilai.Caption = "[1/5] Memproses ID Subjek Pajak: " & Round(i / pNilai.Max * 100, 0) & "%"
        LNilai.Refresh
        LNilai.Visible = False
        pNilai.Value = i
        'Data1(i) = Trim(rPajak!SUBJEK_PAJAK_ID) '& "-" & rPajak!JNS_Bumi
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!SUBJEK_PAJAK_ID)
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", "-"
        'vBangunan.Refresh
            vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", "-"
            vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", "-"
            vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", "-"
            vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
            vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
            vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
            vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", Format(Now, "yyyy")
            'vBangunan.ListItems.Item(i).ListSubItems.Add 11, "", rPajak!NJOP_BUMI * 1 + rPajak!NJOP_BNG * 1
            vBangunan.Refresh
        rPajak.MoveNext
        xPro(0).Text = i
        xPro(0).Refresh
    Loop
    
'n_STR = "select subjek_pajak_id,jns_bumi from QOBJEKPAJAK WHERE JNS_BUMI='1' Group by SUBJEK_PAJAK_ID,JNS_BUMI ORDER BY SUBJEK_PAJAK_ID ASC"
n_STR = "SELECT * FROM QOBJEKPAJAK ORDER BY NJOP_BUMI+NJOP_BNG, SUBJEK_PAJAK_ID ASC" 'SUBJEK_PAJAK_ID, KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP, NJOP_BUMI,NJOP_BNG]*1) AS JTotal From QOBJEKPAJAK ORDER BY SUBJEK_PAJAK_ID, JTOTAL DESC"
'n_str = "SELECT QOBJEKPAJAK.SUBJEK_PAJAK_ID, Max([NJOP_BUMI]*1+[NJOP_BNG]*1) AS JTotal, Max(QOBJEKPAJAK.KD_PROPINSI) AS MaxOfKD_PROPINSI, Max(QOBJEKPAJAK.KD_DATI2) AS MaxOfKD_DATI2, Max(QOBJEKPAJAK.KD_Kecamatan) AS MaxOfKD_Kecamatan, Max(QOBJEKPAJAK.KD_KELURAHAN) AS MaxOfKD_KELURAHAN, Max(QOBJEKPAJAK.KD_BLOK) AS MaxOfKD_BLOK, Max(QOBJEKPAJAK.NO_URUT) AS MaxOfNO_URUT, Max(QOBJEKPAJAK.KD_JNS_OP) AS MaxOfKD_JNS_OP From QOBJEKPAJAK GROUP BY QOBJEKPAJAK.SUBJEK_PAJAK_ID ORDER BY QOBJEKPAJAK.SUBJEK_PAJAK_ID, Max([NJOP_BUMI]*1+[NJOP_BNG]*1)"
'n_str = "SELECT QOBJEKPAJAK.SUBJEK_PAJAK_ID, Max([NJOP_BUMI]*1+[NJOP_BNG]*1) AS JTotal, Max(QOBJEKPAJAK.KD_PROPINSI) AS MaxOfKD_PROPINSI, Max(QOBJEKPAJAK.KD_DATI2) AS MaxOfKD_DATI2, Max(QOBJEKPAJAK.KD_Kecamatan) AS MaxOfKD_Kecamatan, Max(QOBJEKPAJAK.KD_KELURAHAN) AS MaxOfKD_KELURAHAN, Max(QOBJEKPAJAK.KD_BLOK) AS MaxOfKD_BLOK, Max(QOBJEKPAJAK.NO_URUT) AS MaxOfNO_URUT, Max(QOBJEKPAJAK.KD_JNS_OP) AS MaxOfKD_JNS_OP, Count(QOBJEKPAJAK.SUBJEK_PAJAK_ID) AS CountOfSUBJEK_PAJAK_ID, Max(QOBJEKPAJAK.JNS_BUMI) AS MaxOfJNS_BUMI From QOBJEKPAJAK GROUP BY QOBJEKPAJAK.SUBJEK_PAJAK_ID ORDER BY QOBJEKPAJAK.SUBJEK_PAJAK_ID, Max([NJOP_BUMI]*1+[NJOP_BNG]*1)"
openDB (n_STR)
tJum.Text = rPajak.RecordCount
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
i = 0: J = 0
pNilai.Max = tJum.Text
pNilai.Min = 1

    Do While Not rPajak.EOF
        'If rPajak!maxofJNS_BUMI * 1 = 1 Or (rPajak!maxofJNS_BUMI * 1 > 1 And rPajak!CountofSUBJEK_PAJAK_ID * 1 > 1) Then
        i = i + 1
        
'        For i = 1 To rPajak.RecordCount ' vBangunan.ListItems.Count
'            If Trim(vBangunan.ListItems.Item(i).ListSubItems(1).Text) = Trim(rPajak!subjek_pajak_id) Then
            J = J + 1
            If J > 4 Then J = 1
            LNilai.Visible = True
            LNilai.Caption = "[2/5] Proses Pemilihan NJOP Terbesar: " & Round(i / pNilai.Max * 100, 0) & "%"
            LNilai.Refresh
            LNilai.Visible = False
            pNilai.Value = i
'        If Data1(i) = Trim(rPajak!subjek_pajak_id) Then
'            Data2(i) = rPajak!KD_PROPINSI
'            Data3(i) = rPajak!KD_DATI2
'            Data4(i) = rPajak!KD_KECAMATAN
'            Data5(i) = rPajak!KD_KELURAHAN
'            Data6(i) = rPajak!KD_BLOK
'            Data7(i) = rPajak!NO_URUT
'            Data8(i) = rPajak!KD_JNS_OP
            'Data2(i) = ccTahun.Text
           

'            vBangunan1.ListItems.Add i, "", Format(i, "#")
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!SUBJEK_PAJAK_ID)
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 3, "", rPajak!JTotal
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 3, "", rPajak!maxofKD_PROPINSI
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 4, "", rPajak!maxofKD_DATI2
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 5, "", rPajak!maxofKD_KECAMATAN
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 6, "", rPajak!maxofKD_KELURAHAN
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 7, "", rPajak!maxofKD_BLOK
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!maxofNO_URUT
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 9, "", rPajak!maxofKD_JNS_OP
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 10, "", rPajak!maxofJNS_BUMI
'            vBangunan1.ListItems.Item(i).ListSubItems.Add 11, "", rPajak!CountofSUBJEK_PAJAK_ID
            
            vBangunan1.ListItems.Add i, "", Format(i, "#")
            vBangunan1.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
            vBangunan1.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!SUBJEK_PAJAK_ID)
            vBangunan1.ListItems.Item(i).ListSubItems.Add 3, "", rPajak!KD_PROPINSI
            vBangunan1.ListItems.Item(i).ListSubItems.Add 4, "", rPajak!KD_DATI2
            vBangunan1.ListItems.Item(i).ListSubItems.Add 5, "", rPajak!KD_KECAMATAN
            vBangunan1.ListItems.Item(i).ListSubItems.Add 6, "", rPajak!KD_KELURAHAN
            vBangunan1.ListItems.Item(i).ListSubItems.Add 7, "", rPajak!KD_BLOK
            vBangunan1.ListItems.Item(i).ListSubItems.Add 8, "", rPajak!NO_URUT
            vBangunan1.ListItems.Item(i).ListSubItems.Add 9, "", rPajak!KD_JNS_OP
            vBangunan1.ListItems.Item(i).ListSubItems.Add 10, "", rPajak!JNS_BUMI
            vBangunan1.ListItems.Item(i).ListSubItems.Add 11, "", rPajak!NJOP_BUMI * 1 + rPajak!NJOP_BNG
            
            'vBangunan1.ListItems.Item(i).ListSubItems.Add 11, "", rPajak!CountofSUBJEK_PAJAK_ID
'            vBangunan.ListItems.Item(i).ListSubItems(3).Text = rPajak!KD_PROPINSI
'            vBangunan.ListItems.Item(i).ListSubItems(4).Text = rPajak!KD_DATI2
'            vBangunan.ListItems.Item(i).ListSubItems(5).Text = rPajak!KD_KECAMATAN
'            vBangunan.ListItems.Item(i).ListSubItems(6).Text = rPajak!KD_KELURAHAN
'            vBangunan.ListItems.Item(i).ListSubItems(7).Text = rPajak!KD_BLOK
'            vBangunan.ListItems.Item(i).ListSubItems(8).Text = rPajak!NO_URUT
'            vBangunan.ListItems.Item(i).ListSubItems(9).Text = rPajak!KD_JNS_OP
'            vBangunan.ListItems.Item(i).ListSubItems(10).Text = ccTahun.Text
        'End If
        xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
        xPro(0).Text = i
        xPro(0).Refresh
        xPro(1).Text = xNOP
        xPro(1).Refresh
'End If
       ' Next
    'End If
        rPajak.MoveNext


    Loop

'
'i = 0: J = 0
'    Titik(1) = "..": Titik(2) = "....": Titik(3) = "......": Titik(4) = "........"
'pNilai.Max = tJum.Text
'pNilai.Min = 1
'
'
'        For i = 1 To vBangunan.ListItems.Count
'            J = J + 1
'            If J > 4 Then J = 1
'            LNilai.Visible = True
'            LNilai.Caption = "Proses Penentuan Jumlah Terbesar" & Titik(J)
'            LNilai.Refresh
'            LNilai.Visible = False
'            pNilai.Value = i
'            For L = 1 To vBangunan1.ListItems.Count
''            If Trim(vBangunan.ListItems.Item(i).ListSubItems(1).Text) = Trim(rPajak!subjek_pajak_id) Then
'
''        If Data1(i) = Trim(rPajak!subjek_pajak_id) Then
''            Data2(i) = rPajak!KD_PROPINSI
''            Data3(i) = rPajak!KD_DATI2
''            Data4(i) = rPajak!KD_KECAMATAN
''            Data5(i) = rPajak!KD_KELURAHAN
''            Data6(i) = rPajak!KD_BLOK
''            Data7(i) = rPajak!NO_URUT
''            Data8(i) = rPajak!KD_JNS_OP
'            'Data2(i) = ccTahun.Text
'
'            If Trim(vBangunan1.ListItems.Item(L).ListSubItems(2).Text) = Trim(vBangunan.ListItems.Item(i).ListSubItems(2).Text) Then
'                vBangunan.ListItems.Item(i).ListSubItems(3).Text = vBangunan1.ListItems.Item(L).ListSubItems(3).Text
'                vBangunan.ListItems.Item(i).ListSubItems(4).Text = vBangunan1.ListItems.Item(L).ListSubItems(4).Text
'                vBangunan.ListItems.Item(i).ListSubItems(5).Text = vBangunan1.ListItems.Item(L).ListSubItems(5).Text
'                vBangunan.ListItems.Item(i).ListSubItems(6).Text = vBangunan1.ListItems.Item(L).ListSubItems(6).Text
'                vBangunan.ListItems.Item(i).ListSubItems(7).Text = vBangunan1.ListItems.Item(L).ListSubItems(7).Text
'                vBangunan.ListItems.Item(i).ListSubItems(8).Text = vBangunan1.ListItems.Item(L).ListSubItems(8).Text
'                vBangunan.ListItems.Item(i).ListSubItems(9).Text = vBangunan1.ListItems.Item(L).ListSubItems(9).Text
'                vBangunan.Refresh
'                'vBangunan.ListItems.Item(i).ListSubItems(10).Text = vBangunan1.ListItems.Item(L).ListSubItems(10).Text
'            End If
'        Next
'        'xNOP = Trim(rPajak![KD_PROPINSI]) & "." & Trim(rPajak![KD_DATI2]) & "." & Trim(rPajak![KD_KECAMATAN]) & "." & Trim(rPajak![KD_KELURAHAN]) & "." & Trim(rPajak![KD_BLOK]) & "-" & Trim(rPajak![NO_URUT]) & "." & Trim(rPajak![KD_JNS_OP])
'        xNOP = vBangunan.ListItems.Item(i).ListSubItems(3).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(4).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(5).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(6).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(7).Text & "-" & vBangunan.ListItems.Item(i).ListSubItems(8).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(9).Text
'        xPro(0).Text = i
'        xPro(0).Refresh
'        xPro(1).Text = xNOP
'        xPro(1).Refresh
''
'       Next


pNilai.Max = tJum.Text
pNilai.Min = 1

    For K = 1 To vBangunan1.ListItems.Count
        
            
            LNilai.Visible = True
            LNilai.Caption = "[3/5] Memilih NOP yang dikenakan NJOPTKP: " & Round(K / pNilai.Max * 100, 0) & "%"
            LNilai.Refresh
            LNilai.Visible = False
            pNilai.Value = K
        For i = 1 To vBangunan.ListItems.Count
            
            'If (vBangunan.ListItems.Item(i).ListSubItems(11).Text * 1 = vBangunan.ListItems.Item(K).ListSubItems(11).Text * 1) And (Trim(vBangunan.ListItems.Item(i).ListSubItems(2).Text) = Trim(vBangunan.ListItems.Item(K).ListSubItems(2).Text)) Then
            If (Trim(vBangunan.ListItems.Item(i).ListSubItems(2).Text) = Trim(vBangunan1.ListItems.Item(K).ListSubItems(2).Text)) Then
            'i = i + 1
            
        'Data1(i) = Trim(rPajak!SUBJEK_PAJAK_ID) '& "-" & rPajak!JNS_Bumi
'        If vBangunan1.ListItems.Item(i).ListSubItems(10).Text * 1 = 1 Or (vBangunan1.ListItems.Item(i).ListSubItems(10).Text * 1 > 1 And vBangunan1.ListItems.Item(i).ListSubItems(11).Text > 1) Then
'        vBangunan.ListItems.Add i, "", Format(i, "#")
'        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
'        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak!SUBJEK_PAJAK_ID)
'        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", "-"
'        'vBangunan.Refresh
'            vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", "-"
'            vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", "-"
'            vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", "-"
'            vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
'            vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
'            vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
'            vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", Format(Now, "yyyy")
            'For K = 1 To vBangunan.ListItems.Count
                'vBangunan.ListItems.Item(i).ListSubItems(1).Text = vBangunan1.ListItems.Item(K).ListSubItems(1).Text
                'vBangunan.ListItems.Item(i).ListSubItems(2).Text = vBangunan1.ListItems.Item(K).ListSubItems(2).Text
                vBangunan.ListItems.Item(i).ListSubItems(3).Text = vBangunan1.ListItems.Item(K).ListSubItems(3).Text
                vBangunan.ListItems.Item(i).ListSubItems(4).Text = vBangunan1.ListItems.Item(K).ListSubItems(4).Text
                vBangunan.ListItems.Item(i).ListSubItems(5).Text = vBangunan1.ListItems.Item(K).ListSubItems(5).Text
                vBangunan.ListItems.Item(i).ListSubItems(6).Text = vBangunan1.ListItems.Item(K).ListSubItems(6).Text
                vBangunan.ListItems.Item(i).ListSubItems(7).Text = vBangunan1.ListItems.Item(K).ListSubItems(7).Text
                vBangunan.ListItems.Item(i).ListSubItems(8).Text = vBangunan1.ListItems.Item(K).ListSubItems(8).Text
                vBangunan.ListItems.Item(i).ListSubItems(9).Text = vBangunan1.ListItems.Item(K).ListSubItems(9).Text
                'vBangunan.Refresh
            'Next
            xPro(1).Text = vBangunan.ListItems.Item(i).ListSubItems(3).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(4).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(5).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(6).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(7).Text & "-" & vBangunan.ListItems.Item(i).ListSubItems(8).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(9).Text
        xPro(1).Refresh
        End If
        
        Next
    xPro(0).Text = K
        xPro(0).Refresh
        
    Next

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
    
    Screen.MousePointer = vbDefault

End Sub
Sub sv_NJOPTKP()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
 LNilai.Visible = True
 LNilai.Caption = "[4/5] Proses Penghapusan NJOPTKP Lama...."
 LNilai.Refresh
's_SQL = "DELETE  From xNJOPTKP"
s_SQL = "DELETE  From DAT_SUBJEK_PAJAK_NJOPTKP"
openDB (s_SQL)
LNilai.Refresh
LNilai.Visible = False
J = 0
pNilai.Max = vBangunan.ListItems.Count
pNilai.Min = 1
'Titik(1) = "..": Titik(2) = "....": Titik(3) = "......": Titik(4) = "........"
For i = 1 To vBangunan.ListItems.Count
        J = J + 1
        If J > 4 Then J = 1
        LNilai.Visible = True
        LNilai.Caption = "[5/5] Proses Pembuatan NJOPTKP Baru..: " & Round(i / pNilai.Max * 100, 0) & "%"
        LNilai.Refresh
        LNilai.Visible = False
        pNilai.Value = i
         's_SQL = "Insert Into DAT_SUBJEK_PAJAK_NJOPTKP values ('" & vBangunan.ListItems.Item(i).ListSubItems(2).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(3).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(4).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(5).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(6).Text & "'," & _
                "'" & vBangunan.ListItems.Item(i).ListSubItems(7).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(8).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(9).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(10).Text & "')"
        s_SQL = "Insert Into DAT_SUBJEK_PAJAK_NJOPTKP values ('" & vBangunan.ListItems.Item(i).ListSubItems(2).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(3).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(4).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(5).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(6).Text & "'," & _
                "'" & vBangunan.ListItems.Item(i).ListSubItems(7).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(8).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(9).Text & "','" & vBangunan.ListItems.Item(i).ListSubItems(10).Text & "')"
        's_SQL = "Insert Into xNJOPTKP values (Data1(i),'" & Data2(i) & "','" & Data3(i) & "','" & Data4(i) & "','" & Data5(i) & "','" & Data6(i) & "','" & Data7(i) & "','" & Data8(i) & "','" & ccTahun.Text & "')"
        openDB (s_SQL)
        xNOP = vBangunan.ListItems.Item(i).ListSubItems(3).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(4).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(5).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(6).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(7).Text & "-" & vBangunan.ListItems.Item(i).ListSubItems(8).Text & "." & vBangunan.ListItems.Item(i).ListSubItems(9).Text
        xPro(0).Text = i
        xPro(0).Refresh
        xPro(1).Text = xNOP 'Data2(i) & "." & Data3(i) & "." & Data4(i) & "." & Data5(i) & "." & Data6(i) & "-" & Data7(i) & "." & Data8(i)
        xPro(1).Refresh

Next
LNilai.Visible = True
LNilai.Caption = "Sukses..!"
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub vBangunan1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan1.SortKey = ColumnHeader.Index - 1
vBangunan1.Sorted = True
vBangunan1.Sorted = False
vBangunan1.SortOrder = lvwAscending
End Sub


