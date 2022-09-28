VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTarif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penentuan Tarif Minimal"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6090
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6225
      TabIndex        =   21
      Top             =   4080
      Width           =   6225
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
      TabIndex        =   20
      Top             =   0
      Width           =   6225
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   375
      Width           =   5835
      Begin VB.CheckBox chPajak 
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
         Left            =   1845
         TabIndex        =   10
         Top             =   270
         Width           =   1695
      End
      Begin VB.CheckBox chPajak 
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
         TabIndex        =   9
         Top             =   270
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chPajak 
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
         Left            =   3780
         TabIndex        =   11
         Top             =   270
         Width           =   1665
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
      Left            =   3435
      TabIndex        =   8
      Top             =   3645
      Width           =   810
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
      Left            =   2640
      TabIndex        =   7
      Top             =   3645
      Width           =   810
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
      Height          =   360
      Left            =   1845
      TabIndex        =   6
      Top             =   3645
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
      Height          =   2625
      Left            =   120
      TabIndex        =   12
      Top             =   930
      Width           =   5835
      Begin VB.TextBox tPetugas 
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
         Left            =   1845
         TabIndex        =   4
         Top             =   1740
         Width           =   2505
      End
      Begin VB.TextBox tNilai 
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
         Left            =   1845
         TabIndex        =   3
         Top             =   1365
         Width           =   2505
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
         Left            =   1845
         TabIndex        =   0
         Text            =   "ccTahun"
         Top             =   255
         Width           =   1380
      End
      Begin VB.TextBox tSK 
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
         Left            =   1845
         TabIndex        =   1
         Top             =   645
         Width           =   3840
      End
      Begin MSComCtl2.DTPicker dSK 
         Height          =   315
         Left            =   1845
         TabIndex        =   2
         Top             =   1005
         Width           =   2520
         _ExtentX        =   4445
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
         Format          =   152174593
         CurrentDate     =   41486
      End
      Begin MSComCtl2.DTPicker dRekam 
         Height          =   315
         Left            =   1860
         TabIndex        =   5
         Top             =   2115
         Width           =   2520
         _ExtentX        =   4445
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
         Format          =   152174593
         CurrentDate     =   41486
      End
      Begin VB.Label Label6 
         Caption         =   "Tanggal"
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
         Left            =   405
         TabIndex        =   18
         Top             =   2145
         Width           =   1320
      End
      Begin VB.Label Label5 
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
         Height          =   270
         Left            =   405
         TabIndex        =   17
         Top             =   1785
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "Nilai PBB Minimal"
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
         Left            =   405
         TabIndex        =   16
         Top             =   1380
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal"
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
         Left            =   405
         TabIndex        =   15
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "No. SK"
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
         Left            =   405
         TabIndex        =   14
         Top             =   690
         Width           =   1320
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
         Height          =   225
         Left            =   405
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTarif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CEK
Private Sub ccTahun_Click()
On Error GoTo Salah
xSQL = "Select * From PBB_MINIMAL where THN_PBB_MINIMAL='" & Trim(ccTahun.Text) & "'order by THN_PBB_MINIMAL asc"
openDB (xSQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then 'Jika Ditemukan
    With rPajak
        If IsNull(!NO_SK_PBB_MINIMAL) = True Then
            tSK.Text = "-"
        Else
            tSK.Text = !NO_SK_PBB_MINIMAL
        End If
        If IsNull(!TGL_SK_PBB_MINIMAL) = True Then
            dSK.Value = "01/01/1900"
        Else
            dSK.Value = Format(!TGL_SK_PBB_MINIMAL, "dd/mm/yyyy")
        End If
        If IsNull(!NILAI_PBB_MINIMAL) = True Then
            tNilai.Text = 0
        Else
            tNilai.Text = !NILAI_PBB_MINIMAL
        End If
        
        If IsNull(!TGL_REKAM_PBB_MINIMAL) = True Then
            dRekam.Value = "01/01/1900"
        Else
            dRekam.Value = Format(!TGL_REKAM_PBB_MINIMAL, "DD/MM/YYYY")
        End If
        
        If IsNull(!TGL_SK_PBB_MINIMAL) = True Then
            tPetugas.Text = 0
        Else
            tPetugas.Text = !NIP_PEREKAM_PBB_MINIMAL
        End If
        
    End With
Else
    Aktif
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

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
            ccTahun_Click
            Exit Sub
        End If
          If i = ccTahun.ListCount - 1 Then
            If UCase(ccTahun.List(i)) Like "*" + UCase(ccTahun.Text) + "*" = False Then
                ccTahun.Text = ccTahun.List(0)
                ccTahun_Click
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub chPajak_Click(Index As Integer)
On Error GoTo Salah

ccTahun.Text = ""
Aktif



Select Case Index
Case 1
    If chPajak(1).Value = 1 Then
        chPajak(2).Value = 0
        chPajak(3).Value = 0
        cmdSave.Caption = "&Save"
        CEK = 1
    End If
    
Case 2
    
   If chPajak(2).Value = 1 Then
    chPajak(1).Value = 0
    chPajak(3).Value = 0
'    xTanya = MsgBox("Apa anda yakin menghapus NAMA JALAN?", vbQuestion + vbYesNo, "Penghapusan")
'    If xTanya = vbYes Then
'    Else
'        chPajak(1).Value = 1
'        chPajak(2).Value = 0
'    End If
'
   cmdSave.Caption = "&Delete"
   CEK = 2
   End If
Case 3
    If chPajak(3).Value = 1 Then
        chPajak(1).Value = 0
        chPajak(2).Value = 0
'        xTanya = MsgBox("Apa anda yakin melakukan pemutakhiran NAMA JALAN?", vbQuestion + vbYesNo, "Pemutakhiran")
'        If xTanya = vbYes Then
'        Else
'            chPajak(1).Value = 1
'            chPajak(3).Value = 0
'        End If
    cmdSave.Caption = "&Update"
    CEK = 3
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
    KeyAscii = 0
End If
End Sub

Private Sub cmdClear_Click()
Aktif
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
    CTANYA = MsgBox("Apa Anda Yakin Menyimpan Nilai Minimal PBB?", vbQuestion + vbYesNo, "Simpan")
    If CTANYA = vbYes Then
        CALL_OPERASI (1)
        Aktif
    End If
Case 2
    CTANYA = MsgBox("Apa Anda Yakin Menghapus Nilai Minimal PBB?", vbQuestion + vbYesNo, "Hapus")
    If CTANYA = vbYes Then
        CALL_OPERASI (2)
        Aktif
    End If
Case 3
    CTANYA = MsgBox("Apa Anda Yakin Mengupdate Nilai Minimal PBB?", vbQuestion + vbYesNo, "Update")
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
cmdSave.Caption = "&Save"
Aktif
End Sub

Private Sub tNilai_GotFocus()
On Error Resume Next
Call c_blok(tNilai)
tNilai.Alignment = 0
End Sub

Private Sub tNilai_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789.,", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub
Sub CALL_OPERASI(CEK1)
On Error GoTo Salah
If ccTahun.Text = "" Then
    MsgBox "Data Belum Lengkap...", vbCritical, "Error"
    Exit Sub
End If
xSQL = "Select * From PBB_MINIMAL where THN_PBB_MINIMAL='" & Trim(ccTahun.Text) & "'order by THN_PBB_MINIMAL asc"
openDB (xSQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst


Select Case CEK1
Case 1
    If Not rPajak.EOF Then 'Jika Ditemukan
        MsgBox "Data Sudah Ada...", vbCritical, "Error"
        Exit Sub
    End If
    rPajak.AddNew
    rPajak!KD_PROPINSI = "12"
    rPajak!KD_DATI2 = "12"
    rPajak!THN_PBB_MINIMAL = ccTahun.Text
    rPajak!NO_SK_PBB_MINIMAL = tSK.Text
    rPajak!TGL_SK_PBB_MINIMAL = Format(dSK.Value, "dd/mm/yyyy")
    rPajak!NILAI_PBB_MINIMAL = tNilai.Text
    rPajak!TGL_REKAM_PBB_MINIMAL = Format(dRekam.Value, "DD/MM/YYYY")
    rPajak!NIP_PEREKAM_PBB_MINIMAL = Trim(tPetugas.Text)
    rPajak.Update
Case 2
    If rPajak.EOF Then 'Jika Data Tidak Ditemukan
        MsgBox "Data Belum Ada!", vbCritical, "Error"
        Exit Sub
    End If
    rPajak.Delete adAffectCurrent
    rPajak.Update
Case 3
    If rPajak.EOF Then 'Jika Data Tidak Ditemukan
        MsgBox "Data Belum Ada!", vbCritical, "Error"
        Exit Sub
    End If
    rPajak!KD_PROPINSI = "12"
    rPajak!KD_DATI2 = "12"
    rPajak!THN_PBB_MINIMAL = ccTahun.Text
    rPajak!NO_SK_PBB_MINIMAL = tSK.Text
    rPajak!TGL_SK_PBB_MINIMAL = Format(dSK.Value, "dd/mm/yyyy")
    rPajak!NILAI_PBB_MINIMAL = tNilai.Text
    rPajak!TGL_REKAM_PBB_MINIMAL = Format(dRekam.Value, "DD/MM/YYYY")
    rPajak!NIP_PEREKAM_PBB_MINIMAL = Trim(tPetugas.Text)
    rPajak.Update
End Select
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub Aktif()
On Error Resume Next
tSK.Text = "-"
tNilai.Text = 0
dSK.Value = Format(Now, "dd/mm/yyyy")
dRekam.Value = Format(Now, "dd/mm/yyyy")
tPetugas.Text = 0
tPetugas.Alignment = 1: tSK.Alignment = 1: tNilai.Alignment = 1
End Sub

Private Sub tNilai_LostFocus()
On Error Resume Next
tNilai.Alignment = 1
End Sub

Private Sub tPetugas_GotFocus()
On Error Resume Next
Call c_blok(tPetugas)
tPetugas.Alignment = 0
End Sub

Private Sub tPetugas_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789.,", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub tPetugas_LostFocus()
On Error Resume Next
tPetugas.Alignment = 1
End Sub

Private Sub tSK_GotFocus()
On Error Resume Next
Call c_blok(tSK)
tSK.Alignment = 0
End Sub

Private Sub tSK_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    SendKeys "{tab}"
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

Private Sub tSK_LostFocus()
On Error Resume Next
tSK.Text = Rep(tSK.Text)
tSK.Alignment = 1
End Sub
