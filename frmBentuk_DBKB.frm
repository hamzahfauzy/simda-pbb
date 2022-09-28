VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBentuk_DBKB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pembentukan DBKB Standard Komponen Utama dan Material Otomatis"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4980
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -195
      ScaleHeight     =   360
      ScaleWidth      =   5385
      TabIndex        =   8
      Top             =   1995
      Width           =   5385
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   135
      TabIndex        =   6
      Top             =   390
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
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
      Height          =   330
      Left            =   2310
      TabIndex        =   2
      Top             =   1545
      Width           =   810
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
      Height          =   330
      Left            =   1515
      TabIndex        =   1
      Top             =   1545
      Width           =   810
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   2760
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
         Left            =   660
         TabIndex        =   0
         Text            =   "ccTahun"
         Top             =   255
         Width           =   1920
      End
      Begin VB.Label Label1 
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
         Left            =   75
         TabIndex        =   4
         Top             =   285
         Width           =   1215
      End
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
      Height          =   750
      Left            =   2880
      TabIndex        =   5
      Top             =   660
      Width           =   1965
      Begin VB.CheckBox chPajak 
         Caption         =   "DBKB Material"
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
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chPajak 
         Caption         =   "DBKB Standard"
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
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Value           =   1  'Checked
         Width           =   1710
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   5385
      TabIndex        =   7
      Top             =   0
      Width           =   5385
   End
End
Attribute VB_Name = "frmBentuk_DBKB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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


Private Sub chPajak_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub




Private Sub cmdExit_Click()
On Error Resume Next
Unload Me
cBentuk = ""
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
ProgressBar1.Visible = True
If cBentuk = 1 Then

    TANYA = MsgBox("Proses DBKB Standard otomatis?", vbInformation + vbYesNo, "Info")
    If TANYA = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
    If chPajak(1).Value = 0 And chPajak(2).Value = 0 Then
        MsgBox "Jenis DBKB Belum Dipilih...!", vbCritical, "Error"
        chPajak(1).Value = 1
        chPajak(2).Value = 1
        Screen.MousePointer = vbDefault: Exit Sub
    End If
    '--Pembentukan DBKB Utama
    If chPajak(1).Value = 1 Then
        For i = 1 To 20
            ProgressBar1.Value = i
        Next
        C_STR1 = "H_SATUAN '" & ccTahun.Text & "' "
        openDB (C_STR1)
        For i = 21 To 40
            ProgressBar1.Value = i
        Next
        C_STR2 = "H_KEGIATAN '" & ccTahun.Text & "'"
        openDB (C_STR2)
        For i = 41 To 60
            ProgressBar1.Value = i
        Next
        C_STR3 = "HITUNG_DBKB_STANDARD '" & ccTahun.Text & "'"
        openDB (C_STR3)
        For i = 61 To 80
            ProgressBar1.Value = i
        Next
        C_STR4 = "HITUNG_DBKB_FINAL '" & ccTahun.Text & "'"
        openDB (C_STR4)
        For i = 81 To 90
            ProgressBar1.Value = i
        Next
    End If
    '---Pembentukan DBKB Material
    If chPajak(2).Value = 1 Then
        C_STR5 = "DBKB_MAT_HARGA_SATUAN '" & ccTahun.Text & "'"
        openDB (C_STR5)
        C_STR6 = "DBKB_MAT_SEBELUM_ADJUSTMENT '" & ccTahun.Text & "'"
        openDB (C_STR6)
        C_STR7 = "DBKB_MAT_ADJUSTMENT '" & ccTahun.Text & "'"
        openDB (C_STR7)
    End If
        For i = 91 To 100
            ProgressBar1.Value = i
        Next
ElseIf cBentuk = 2 Then
    If chPajak(1).Value = 1 Then
        For i = 1 To 25
            ProgressBar1.Value = i
        Next
        CC_STR1 = "HITUNG_HARGA_KEGIATAN_JPB8 '" & ccTahun.Text & "'"
        openDB (CC_STR1)
        For i = 26 To 50
            ProgressBar1.Value = i
        Next
        CC_STR2 = "HITUNG_DBKB_JPB8 '" & ccTahun.Text & "'"
        openDB (CC_STR2)
        For i = 51 To 75
            ProgressBar1.Value = i
        Next
        CC_STR3 = "HITUNG_DBKB_JPB8_STLH_ADJ '" & ccTahun.Text & "'"
        openDB (CC_STR3)
        For i = 76 To 100
            ProgressBar1.Value = i
        Next
    End If
    If chPajak(2).Value = 1 Then
        For i = 1 To 75
            ProgressBar1.Value = i
        Next
        CC_STR4 = "HITUNG_DBKB_JPB3 '" & ccTahun.Text & "'"
        openDB (CC_STR4)
        For i = 76 To 100
            ProgressBar1.Value = i
        Next
    End If
End If
MsgBox "SUKSES...!"
ProgressBar1.Visible = False
dbPajak.Close
Set dbPajak = Nothing
Set rPajak = Nothing
Screen.MousePointer = vbDefault
End Sub

Private Sub dRekam_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
End If
End Sub

Private Sub dSK_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
End If
End Sub


Private Sub Form_Activate()
On Error Resume Next
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
ccTahun.Clear
ccTahun.Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    ccTahun.AddItem i
Next
If cBentuk = 1 Then
    chPajak(1).Caption = "DBKB Standard"
    chPajak(2).Caption = "DBKB Material"
    Me.Caption = "DBKB Standard: Umum dan Material"
ElseIf cBentuk = 2 Then
    chPajak(1).Caption = "DBKB JPB8"
    chPajak(2).Caption = "DBKB JPB3"
    Me.Caption = "DBKB Non Standard: JPB3_JPB8"
End If
ProgressBar1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
cBentuk = ""
End Sub


