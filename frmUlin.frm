VERSION 5.00
Begin VB.Form frmUlin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status Kayu Ulin Sebagai Pembentuk Dominan Bangunan Kayu?"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6045
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
      Left            =   105
      TabIndex        =   10
      Top             =   45
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   6
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
      Left            =   2730
      TabIndex        =   3
      Top             =   1800
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
      Left            =   1935
      TabIndex        =   2
      Top             =   1800
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
      Height          =   1080
      Left            =   105
      TabIndex        =   7
      Top             =   555
      Width           =   5835
      Begin VB.ComboBox cUlin 
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
         Left            =   4215
         TabIndex        =   1
         Text            =   "ccTahun"
         Top             =   390
         Width           =   1380
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
         Left            =   1560
         TabIndex        =   0
         Text            =   "ccTahun"
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label Label3 
         Caption         =   "Kayu Ulin"
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
         Left            =   3405
         TabIndex        =   9
         Top             =   450
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
         TabIndex        =   8
         Top             =   450
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmUlin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CEK
Private Sub ccTahun_Click()
On Error GoTo Salah
xSQL = "Select * From KAYU_ULIN where THN_STATUS_KAYU_ULIN='" & Trim(ccTahun.Text) & "'order by THN_STATUS_KAYU_ULIN asc"
openDB (xSQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then 'Jika Ditemukan
    If rPajak!STATUS_KAYU_ULIN = "0" Then
        cUlin.Text = cUlin.List(0)
    Else
        cUlin.Text = cUlin.List(1)
    End If
Else
    cUlin.Text = cUlin.List(0)
    
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
End If
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
    CTANYA = MsgBox("Simpan status penggunaan kayu ulin ?", vbQuestion + vbYesNo, "Simpan")
    If CTANYA = vbYes Then
        CALL_OPERASI (1)
        
    End If
Case 2
    CTANYA = MsgBox("Hapus status penggunaan kayu ulin ?", vbQuestion + vbYesNo, "Hapus")
    If CTANYA = vbYes Then
        CALL_OPERASI (2)
        
    End If
Case 3
    CTANYA = MsgBox("Edit status penggunaan kayu ulin ?", vbQuestion + vbYesNo, "Update")
    If CTANYA = vbYes Then
        CALL_OPERASI (3)
        
    End If
End Select
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub cUlin_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub cUlin_LostFocus()
On Error Resume Next
For i = 0 To cUlin.ListCount - 1
        If (UCase(cUlin.List(i)) Like "*" + UCase(cUlin.Text) + "*" = True) Then
            cUlin.Text = cUlin.List(i)
            Exit Sub
        End If
          If i = cUlin.ListCount - 1 Then
            If UCase(cUlin.List(i)) Like "*" + UCase(cUlin.Text) + "*" = False Then
                cUlin.Text = cUlin.List(0)
                Exit Sub
            End If
        End If
    Next
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
'ccTahun.Text = Format(Now, "yyyy")
ccTahun.Text = tck_ulin
For i = Format(Now, "yyyy") To 1900 Step -1
    ccTahun.AddItem i
Next
cUlin.Clear
'cUlin.Text = "0-NO"

cUlin.AddItem "0-NO"
cUlin.AddItem "1-YES"
If ck_Ulin = 0 Then
    cUlin.Text = cUlin.List(0)
Else
    cUlin.Text = cUlin.List(1)
End If
cmdSave.Caption = "&Save"
End Sub


Sub CALL_OPERASI(CEK1)
On Error GoTo Salah
If ccTahun.Text = "" Or cUlin.Text = "" Then
    MsgBox "Data Belum Lengkap...", vbCritical, "Error"
    Exit Sub
End If
xSQL = "Select * From KAYU_ULIN where THN_STATUS_KAYU_ULIN='" & Trim(ccTahun.Text) & "'order by THN_STATUS_KAYU_ULIN asc"
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
    rPajak!THN_STATUS_KAYU_ULIN = ccTahun.Text
    rPajak!STATUS_KAYU_ULIN = Left(Trim(cUlin.Text), 1)
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
    rPajak!THN_STATUS_KAYU_ULIN = ccTahun.Text
    rPajak!STATUS_KAYU_ULIN = Left(Trim(cUlin.Text), 1)
    rPajak.Update
End Select
ck_Ulin = Left(Trim(cUlin.Text), 1)
tck_ulin = ccTahun.Text
ccTahun.Text = ccTahun.List(0)
cUlin.Text = cUlin.List(0)

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
