VERSION 5.00
Begin VB.Form frmSubjek_Pajak 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Subjek Pajak "
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6135
   Icon            =   "frmSubjek_Pajak.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
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
      Height          =   615
      Left            =   0
      TabIndex        =   29
      Top             =   -90
      Width           =   6120
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
         Index           =   3
         Left            =   3870
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
         Left            =   105
         TabIndex        =   0
         Top             =   270
         Value           =   1  'Checked
         Width           =   1680
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
         Index           =   2
         Left            =   1860
         TabIndex        =   1
         Top             =   270
         Width           =   1695
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
      Left            =   3465
      TabIndex        =   16
      Top             =   3975
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
      Left            =   2490
      TabIndex        =   15
      Top             =   3975
      Width           =   990
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
      Left            =   1515
      TabIndex        =   14
      Top             =   3975
      Width           =   990
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3420
      Left            =   0
      TabIndex        =   18
      Top             =   450
      Width           =   6120
      Begin VB.TextBox tID 
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
         Left            =   1605
         TabIndex        =   13
         Text            =   "22272"
         Top             =   3015
         Width           =   1575
      End
      Begin VB.CommandButton cmdSubjek 
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
         Left            =   5490
         TabIndex        =   17
         Top             =   210
         Width           =   315
      End
      Begin VB.ComboBox ccKerja 
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
         ItemData        =   "frmSubjek_Pajak.frx":1CCA
         Left            =   1605
         List            =   "frmSubjek_Pajak.frx":1CDD
         TabIndex        =   5
         Top             =   915
         Width           =   2880
      End
      Begin VB.TextBox tID 
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
         Left            =   1605
         TabIndex        =   12
         Text            =   "PAKPAK BHARAT"
         Top             =   2670
         Width           =   4200
      End
      Begin VB.TextBox tID 
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
         Index           =   7
         Left            =   5055
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "00"
         Top             =   2325
         Width           =   750
      End
      Begin VB.TextBox tID 
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
         Left            =   3825
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "00"
         Top             =   2325
         Width           =   765
      End
      Begin VB.TextBox tID 
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
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   9
         Text            =   "00"
         Top             =   2325
         Width           =   750
      End
      Begin VB.TextBox tID 
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
         Left            =   1605
         TabIndex        =   8
         Top             =   1965
         Width           =   4200
      End
      Begin VB.TextBox tID 
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
         Left            =   1605
         TabIndex        =   7
         Top             =   1620
         Width           =   4200
      End
      Begin VB.TextBox tID 
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
         Index           =   2
         Left            =   1605
         TabIndex        =   6
         Top             =   1275
         Width           =   2850
      End
      Begin VB.TextBox tID 
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
         Left            =   1605
         TabIndex        =   4
         Top             =   570
         Width           =   4200
      End
      Begin VB.TextBox tID 
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
         Left            =   1605
         TabIndex        =   3
         Top             =   225
         Width           =   3870
      End
      Begin VB.Label LID 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Left            =   4830
         TabIndex        =   31
         Top             =   1140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "BLOK/KAV"
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
         Left            =   1635
         TabIndex        =   30
         Top             =   2370
         Width           =   720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         TabIndex        =   28
         Top             =   570
         Width           =   1035
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos"
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
         TabIndex        =   27
         Top             =   3075
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " NPWP"
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
         TabIndex        =   26
         Top             =   1290
         Width           =   810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pekerjaan"
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
         TabIndex        =   25
         Top             =   945
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         Left            =   195
         TabIndex        =   24
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   180
         TabIndex        =   23
         Top             =   1965
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3405
         TabIndex        =   22
         Top             =   2370
         Width           =   255
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4725
         TabIndex        =   21
         Top             =   2370
         Width           =   195
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Kabupaten/Kota"
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
         TabIndex        =   20
         Top             =   2700
         Width           =   1230
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor ID"
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
         Index           =   0
         Left            =   165
         TabIndex        =   19
         Top             =   255
         Width           =   1035
      End
   End
   Begin VB.Image Image1 
      Height          =   9180
      Left            =   0
      Picture         =   "frmSubjek_Pajak.frx":1D1B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "frmSubjek_Pajak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CEK
Private Sub ccKerja_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub ccKerja_LostFocus()
On Error Resume Next
For i = 0 To ccKerja.ListCount - 1
        If (UCase(ccKerja.List(i)) Like "*" + UCase(ccKerja.Text) + "*" = True) Then
            ccKerja.Text = ccKerja.List(i)
            Exit Sub
        End If
          If i = ccKerja.ListCount - 1 Then
            If UCase(ccKerja.List(i)) Like "*" + UCase(ccKerja.Text) + "*" = False Then
                ccKerja.Text = ccKerja.List(0)
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub chPajak_Click(Index As Integer)
On Error GoTo Salah
Select Case Index
Case 1
    If chPajak(1).Value = 1 Then
        chPajak(2).Value = 0
        chPajak(3).Value = 0
        cmdSave.Caption = "&Save"
        cmdSubjek.Enabled = False
        CEK = 1
    Else
        If chPajak(2).Value = 0 And chPajak(3).Value = 0 Then
            chPajak(1).Value = 1
        End If
    End If
    
Case 2
    
   If chPajak(2).Value = 1 Then
    chPajak(1).Value = 0
    chPajak(3).Value = 0
    cmdSubjek.Enabled = True
    cmdSave.Caption = "&Delete"
    CEK = 2
    xTanya = MsgBox("Apa anda yakin menghapus Subjek Pajak?", vbQuestion + vbYesNo, "Penghapusan")
    If xTanya = vbYes Then
        cmdSave.Caption = "&Delete"
    Else
        chPajak(1).Value = 1
        chPajak(2).Value = 0
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
        cmdSubjek.Enabled = True
        cmdSave.Caption = "&Update"
        CEK = 3
        xTanya = MsgBox("Apa anda yakin melakukan pemutakhiran Subjek Pajak?", vbQuestion + vbYesNo, "Pemutakhiran")
        If xTanya = vbYes Then
            cmdSave.Caption = "&Update"
        Else
            chPajak(1).Value = 1
            chPajak(3).Value = 0
        End If
     Else
        If chPajak(1).Value = 0 And chPajak(2).Value = 0 Then
            chPajak(3).Value = 1
        End If
    End If
End Select
tID(0).SetFocus
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub chPajak_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
Aktif
tID(0).SetFocus
LID.Caption = ""
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo Salah
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
    For Each Control In Me
    If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
        If Control.Text = "" Then
            MsgBox "Data masih kosong!", vbCritical, "Tetnong"
            Exit Sub
        End If
    End If
    Next

    CTANYA = MsgBox("Apa Anda Yakin Menyimpan Data Subjek Pajak?", vbQuestion + vbYesNo, "Simpan")
    If CTANYA = vbYes Then
        CALL_OPERASI (1)
        Aktif
    End If
Case 2
    CTANYA = MsgBox("Apa Anda Yakin Menghapus Data Subjek Pajak?", vbQuestion + vbYesNo, "Hapus")
    If CTANYA = vbYes Then
        CALL_OPERASI (2)
        Aktif
    End If
Case 3
    For Each Control In Me
    If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
        If Control.Text = "" Then
            MsgBox "Data masih kosong!", vbCritical, "Tetnong"
            Exit Sub
        End If
    End If
Next

    CTANYA = MsgBox("Apa Anda Yakin Mengupdate Data Subjek Pajak?", vbQuestion + vbYesNo, "Update")
    If CTANYA = vbYes Then
        CALL_OPERASI (3)
        Aktif
    End If
End Select
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cmdSubjek_Click()
On Error Resume Next
'If CEK = 2 Or CEK = 3 Then
    xSQL = "Select * From DAT_SUBJEK_PAJAK where SUBJEK_PAJAK_ID='" & Trim(tID(0).Text) & "'order by SUBJEK_PAJAK_ID asc"
    openDB (xSQL)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If Not rPajak.EOF Then 'Jika Ditemukan
        TAMPIL
        LID.Caption = tID(0).Text
        Exit Sub
    End If
    If CEK = 3 Then
        xSQL = "Select * From DAT_SUBJEK_PAJAK where SUBJEK_PAJAK_ID='" & Trim(LID.Caption) & "'order by SUBJEK_PAJAK_ID asc"
        openDB (xSQL)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
            If Not rPajak.EOF Then 'Jika Ditemukan
            'TAMPIL
            'LID.Caption = tID(0).Text
            Exit Sub
        End If
    End If
    If CEK = 2 Or CEK = 3 Or tID(0).Text = "" Then
    xID = 2
    frmList_Subjek.Show
    End If
'End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2

If xID = 2 Or chPajak(2).Value = 1 Or chPajak(3).Value = 1 Then
    cmdSubjek.Enabled = True
Else
    cmdSubjek.Enabled = False
End If
End Sub

Sub CALL_OPERASI(CEK1)
On Error Resume Next
Select Case CEK1
Case 1 'SAVE
    xSQL = "Select * From DAT_SUBJEK_PAJAK where SUBJEK_PAJAK_ID='" & tID(0).Text & "'order by SUBJEK_PAJAK_ID asc"
    openDB (xSQL)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If Not rPajak.EOF Then 'Jika Ditemukan
        MsgBox "Data Sudah Ada...", vbCritical, "Error"
        Exit Sub
    End If
    rPajak.AddNew
    rPajak!SUBJEK_PAJAK_ID = Trim(tID(0).Text)
    rPajak!Nm_wp = tID(1).Text
    rPajak!JALAN_WP = tID(3).Text
    rPajak!BLOK_KAV_NO_WP = tID(5).Text
    rPajak!RW_WP = tID(6).Text
    rPajak!RT_WP = tID(7).Text
    rPajak!KELURAHAN_WP = tID(4).Text
    rPajak!KOTA_WP = tID(8).Text
    rPajak!KD_POS_WP = tID(9).Text
    rPajak!TELP_WP = "-"
    rPajak!NPWP = tID(2).Text
    rPajak!STATUS_PEKERJAAN_WP = Left(Trim(ccKerja.Text), 2) * 1
    rPajak.Update
Case 2 'HAPUS
    iSQL = "Select SUBJEK_PAJAK_ID From DAT_OBJEK_PAJAK where (SUBJEK_PAJAK_ID)='" & LID.Caption & "'order by SUBJEK_PAJAK_ID asc"
    openDB (iSQL)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        MsgBox "Subjek tidak dapat dihapus, ID Masih Digunakan,," & _
            vbCrLf & "Silahkan hapus data data bangunan dan bumi terlebih dahulu..", vbCritical, "Tetnong..."
        Exit Sub
    rPajak.MoveNext
    Loop
    'uSQL = "Select SUBJEK_PAJAK_ID From DAT_OBJEK_PAJAK where trim(SUBJEK_PAJAK_ID)='" & Trim(LID.Caption) & "'order by SUBJEK_PAJAK_ID asc"
    uSQL = "Select SUBJEK_PAJAK_ID From DAT_SUBJEK_PAJAK where SUBJEK_PAJAK_ID='" & tID(0).Text & "'order by SUBJEK_PAJAK_ID asc"
    openDB (uSQL)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    If rPajak.EOF Then 'Jika Tidak Ditemukan
        MsgBox "Tidak ada data yang akan terhapus...", vbCritical, "Error"
        Exit Sub
    End If
    rPajak.Delete adAffectCurrent
    rPajak.Update
Case 3 'UPDATE
'    tID(0).Text = rPajak!SUBJEK_PAJAK_ID
'    tID(1).Text = rPajak!NM_WP
'    tID(3).Text = rPajak!JALAN_WP
'    tID(5).Text = rPajak!BLOK_KAV_NO_WP
'    tID(6).Text = rPajak!RW_WP
'    tID(7).Text = rPajak!RT_WP
'    tID(4).Text = rPajak!KELURAHAN_WP
'    tID(8).Text = rPajak!KOTA_WP
'    tID(9).Text = rPajak!KD_POS_WP
'    'rPajak!TELP_WP = "-"
'    tID(2).Text = rPajak!NPWP
'    ccKerja.Text = ccKerja.List(rPajak!STATUS_PEKERJAAN_WP * 1)
    'xSQL = "Select * From DAT_SUBJEK_PAJAK where SUBJEK_PAJAK_ID='" & Trim(LID.Caption) & "'order by SUBJEK_PAJAK_ID asc"
    xSQL = "Select * From DAT_SUBJEK_PAJAK where SUBJEK_PAJAK_ID='" & Trim(tID(0).Text) & "'order by SUBJEK_PAJAK_ID asc"
    openDB (xSQL)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst

    If Not rPajak.EOF Then 'Jika Ditemukan
        If Trim(tID(0).Text) <> Trim(LID.Caption) Then
            MsgBox "No ID sudah ada ...!", vbCritical, "Error"
            Exit Sub
        End If
        
    End If

    xSQL = "Select * From DAT_SUBJEK_PAJAK where SUBJEK_PAJAK_ID='" & Trim(LID.Caption) & "'order by SUBJEK_PAJAK_ID asc"
    openDB (xSQL)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst

    If rPajak.EOF Then 'Jika tidak Ditemukan
        MsgBox "Data tidak ditemukan...!", vbCritical, "Error"
        Exit Sub
    End If

    rPajak!SUBJEK_PAJAK_ID = Trim(tID(0).Text)
    rPajak!Nm_wp = tID(1).Text
    rPajak!JALAN_WP = tID(3).Text
    rPajak!BLOK_KAV_NO_WP = tID(5).Text
    rPajak!RW_WP = tID(6).Text
    rPajak!RT_WP = tID(7).Text
    rPajak!KELURAHAN_WP = tID(4).Text
    rPajak!KOTA_WP = tID(8).Text
    rPajak!KD_POS_WP = tID(9).Text
    rPajak!TELP_WP = "-"
    rPajak!NPWP = tID(2).Text
    rPajak!STATUS_PEKERJAAN_WP = Left(Trim(ccKerja.Text), 2) * 1
    rPajak.Update
    uSQL = "Select SUBJEK_PAJAK_ID From DAT_OBJEK_PAJAK where (SUBJEK_PAJAK_ID)='" & Trim(LID.Caption) & "'order by SUBJEK_PAJAK_ID asc"
    openDB (uSQL)
    If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    Do While Not rPajak.EOF
        rPajak!SUBJEK_PAJAK_ID = Trim(tID(0).Text)
        rPajak.Update
    rPajak.MoveNext
    Loop
End Select
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'Salah:
'If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
'MsgBox Err.Number & ": " & Err.Description

End Sub

Sub Aktif()
On Error Resume Next
tID(0).Text = ""
    tID(1).Text = "-"
    tID(3).Text = "-"
    tID(5).Text = "00"
    tID(6).Text = "00"
    tID(7).Text = "00"
    tID(4).Text = "-"
    tID(8).Text = "PAKPAK BHARAT"
    tID(9).Text = "22272"
    'rPajak!TELP_WP = "-"
    tID(2).Text = "-"
    LID.Caption = ""
    
    ccKerja.Text = ccKerja.List(0)
End Sub

Private Sub tID_GotFocus(Index As Integer)
On Error Resume Next
Select Case Index
Case Index
    Call c_blok(tID(Index))
End Select
End Sub

Private Sub tID_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
        
        SendKeys "{tab}"
        KeyAscii = 0
End If
Select Case Index
Case 2, 5, 6, 7, 9
    If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
    End If
End Select
End Sub
Sub c_blok(nControl As TextBox)
On Error Resume Next
nControl.SelStart = 0
nControl.SelLength = Len(nControl.Text)
'nControl.SetFocus
'nControl.Alignment = 0
End Sub

Sub TAMPIL()
On Error Resume Next
tID(0).Text = rPajak!SUBJEK_PAJAK_ID
    tID(1).Text = rPajak!Nm_wp
    tID(3).Text = rPajak!JALAN_WP
    tID(5).Text = rPajak!BLOK_KAV_NO_WP
    tID(6).Text = rPajak!RW_WP
    tID(7).Text = rPajak!RT_WP
    tID(4).Text = rPajak!KELURAHAN_WP
    tID(8).Text = rPajak!KOTA_WP
    tID(9).Text = rPajak!KD_POS_WP
    'rPajak!TELP_WP = "-"
    tID(2).Text = rPajak!NPWP
    ccKerja.Text = ccKerja.List(rPajak!STATUS_PEKERJAAN_WP * 1 - 1)
End Sub

Private Sub tID_LostFocus(Index As Integer)
On Error Resume Next
Select Case Index
    Case Index
        tID(Index).Text = Rep(tID(Index).Text)
    Case 0
       ' If CEK = 1 Or CEK = "" Or CEK = 2 Or (CEK = 3 And UCase(LID.Caption) = UCase(tID(0).Text)) Then
            cmdSubjek_Click
      '  End If
      Case 5, 6, 7
        If tID(Index).Text = "" Then tID(Index).Text = "00"
End Select
End Sub
