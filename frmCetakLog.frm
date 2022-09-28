VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCetakLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Log Data"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   ControlBox      =   0   'False
   Icon            =   "frmCetakLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5925
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      Height          =   510
      Left            =   1770
      TabIndex        =   16
      Top             =   -75
      Width           =   4065
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
         Left            =   3555
         TabIndex        =   18
         Top             =   165
         Visible         =   0   'False
         Width           =   345
      End
      Begin MSMask.MaskEdBox aNOP 
         Height          =   315
         Left            =   510
         TabIndex        =   1
         Top             =   150
         Width           =   3405
         _ExtentX        =   6006
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
         Left            =   495
         TabIndex        =   17
         Top             =   150
         Width           =   3000
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
         Left            =   120
         TabIndex        =   19
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   -30
      ScaleHeight     =   480
      ScaleWidth      =   6000
      TabIndex        =   14
      Top             =   -30
      Width           =   6000
      Begin VB.CheckBox hTunggal 
         BackColor       =   &H80000002&
         Caption         =   "Print by N.O.P"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   1875
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
      Left            =   3480
      TabIndex        =   9
      Top             =   2310
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
      Left            =   2580
      TabIndex        =   8
      Top             =   2310
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Cetak"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   2310
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   -120
      ScaleHeight     =   765
      ScaleWidth      =   6150
      TabIndex        =   15
      Top             =   2040
      Width           =   6150
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1650
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   5985
      Begin VB.CheckBox cLog 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tanggal Log"
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
         Left            =   135
         TabIndex        =   21
         Top             =   975
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dRekam2 
         Height          =   315
         Left            =   3735
         TabIndex        =   6
         Top             =   975
         Width           =   1995
         _ExtentX        =   3519
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
         Format          =   186580993
         CurrentDate     =   41486
      End
      Begin MSComCtl2.DTPicker dRekam1 
         Height          =   315
         Left            =   1485
         TabIndex        =   5
         Top             =   975
         Width           =   1830
         _ExtentX        =   3228
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
         Format          =   186580993
         CurrentDate     =   41486
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
         Left            =   1410
         TabIndex        =   2
         Top             =   -15
         Visible         =   0   'False
         Width           =   1350
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
         Left            =   1485
         TabIndex        =   4
         Top             =   630
         Width           =   4260
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
         Left            =   1485
         TabIndex        =   3
         Top             =   300
         Width           =   4260
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "s.d"
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
         Left            =   3405
         TabIndex        =   20
         Top             =   1050
         Width           =   285
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
         Height          =   165
         Left            =   60
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label8 
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
         Left            =   135
         TabIndex        =   12
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
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
         Left            =   135
         TabIndex        =   11
         Top             =   345
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCetakLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xTT, xTB, QQ
Dim NMIN, cTarif
'Dim xMIN(2), xMAX(2)
Dim xTarif(2)
Dim cMin(2), cMax(2), cTKP(2)
Dim totChar


Private Sub ccKel_Click()
On Error Resume Next
C_KEC = Left(Trim(ccKec.Text), 3)
    C_KEL = Left(Trim(ccKel.Text), 3)
    
End Sub

Private Sub ccTahun_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub



Private Sub cLog_Click()
On Error Resume Next
If cLog.Value = 1 Then dRekam1.Enabled = True: dRekam2.Enabled = True Else dRekam1.Enabled = False: dRekam2.Enabled = False

End Sub

Private Sub cmdCear_Click()
On Error Resume Next
ccTahun.Text = ccTahun.List(0)
ccKec.Text = ""
ccKel.Text = ""
dRekam1.Value = Format(Now, "dd/mm/yyyy")
dRekam2.Value = Format(Now, "dd/mm/yyyy")

hTunggal.Value = 0
Frame4.Visible = False


End Sub

Private Sub cmdExit_Click()
On Error Resume Next
Unload Me
c_Ganti = 0
xID = ""
End Sub


Private Sub cmdNOP1_Click()
On Error GoTo Salah
J_Karakter
If Len(Trim(tNOP(0).Text)) - (totChar * 1) = 24 Then
'    call_data
Else
    xID = 6
    frmLIST_Objek1.Show
End If
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cmdOK_Click()
On Error GoTo Salah
Dim Pesan
Screen.MousePointer = vbHourglass
If ccKec.Text = "" Then ccKec.Text = "*.*"
If ccKel.Text = "" Then ccKel.Text = "*.*"
    C_KEC = Left(Trim(ccKec.Text), 3)
    C_KEL = Left(Trim(ccKel.Text), 3)
    C_TAHUN = ccTahun.Text
    c_NOP = aNOP.Text
    
If J_CETAK = 111 Or J_CETAK = 112 Or J_CETAK = 113 Then
        If hTunggal.Value = 1 Then
            Pesan = "Apa anda yakin cetak log secara tunggal?"
                CetakQ = 1
        Else
            Pesan = "Apa anda yakin cetak log secara massal?"
            If C_KEC = "*.*" And C_KEL = "*.*" Then
                If cLog.Value = 1 Then
                    CetakQ = 2
                Else
                    CetakQ = 3
                End If
                
            ElseIf C_KEC <> "*.*" And C_KEL = "*.*" Then
                If cLog.Value = 1 Then
                    CetakQ = 4
                Else
                    CetakQ = 5
                End If
            Else
                If cLog.Value = 1 Then
                    CetakQ = 6
                Else
                    CetakQ = 7
                End If
            End If
        End If
End If
If J_CETAK <> 113 Then rptPBB.Show
If cmdOK.Caption = "&Hapus" Then cHapus_Log
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub dJTempo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If

End Sub



Private Sub cUrut_Click()
On Error Resume Next
If cUrut.Value = 1 Then
    cRekam.Value = 0
    tSPPT.Visible = True
    tSPPT2.Visible = True
    Label6.Visible = True
    Label10.Visible = True
    Label10.Caption = "[KDBlok].[NoUrut]"
    ccKel.Text = ""
    dRekam1.Visible = False
    dRekam2.Visible = False
Else
    'cRekam.Value = 1
    tSPPT.Text = 0
    tSPPT2.Text = 0
    tSPPT.Visible = False
    tSPPT2.Visible = False
    Label6.Visible = False
    Label10.Visible = False
    
End If
End Sub

Private Sub cUrut_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub dRekam2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If

End Sub

Private Sub dRekam1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If
End Sub

'Private Sub dRekam2_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    SendKeys "{Tab}"
'End If
'End Sub
'
'Private Sub dRekam1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    SendKeys "{Tab}"
'End If
'
'End Sub

Private Sub Form_Activate()
On Error GoTo Salah
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
If J_CETAK = 113 Then cmdOK.Caption = "&Hapus": hTunggal.Caption = "Delete by N.O.P" Else cmdOK.Caption = "&Cetak": hTunggal.Caption = "Print by N.O.P"
SKRG = Format(Now, "YYYY")
dRekam1.Value = "01/01/" & SKRG
dRekam2.Value = "31/12/" & SKRG
If cLog.Value = 1 Then dRekam1.Enabled = True: dRekam2.Enabled = True Else dRekam1.Enabled = False: dRekam2.Enabled = False
If c_Ganti = "" Or c_Ganti = 0 Then
ccTahun.Text = Format(Now, "yyyy")
ccTahun.Clear
ccTahun.Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    ccTahun.AddItem i
Next
If J_CETAK = 111 Then
    Me.Caption = "Print Data Perubahan Objek Pajak"
ElseIf J_CETAK = 112 Then
    Me.Caption = "Print Perubahan Ketetapan Objek Pajak Setelah Perubahan"
Else
    Me.Caption = "Hapus Log Data"
End If

If xID = "" Then
    hTunggal.Value = 0
    Frame4.Visible = False
    
'Else
 '   hTunggal.Value = 1
  '  Frame4.Visible = True
End If
If ccKec.ListCount <= 0 Then
    CALL_KEC
End If
'Frame4.Visible = False
'hTunggal.Value = 0

End If


If hTunggal.Value = 0 Then Frame4.Visible = False
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
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
        ccKec.AddItem "*.*"
    
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
        ccKel.AddItem "*.*"
    
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub

Private Sub ccKec_Click()
On Error GoTo Salah
If ccKec.Text = "*.*" Then
    ccKel.Enabled = False
    ccKel.Text = "*.*"

Else
    ccKel.Enabled = True
    CALL_KEL
End If

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub ccKec_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

If InStr("0123456789*.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
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

If InStr("0123456789*.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
         KeyAscii = 0
        End If
End Sub

Private Sub ccKel_LostFocus()
On Error Resume Next

For i = 0 To ccKel.ListCount - 1
        If (UCase(ccKel.List(i)) Like "*" + UCase(ccKel.Text) + "*" = True) Then
            ccKel.Text = ccKel.List(i)
            ccKel_Click
            Exit Sub
        End If
          If i = ccKel.ListCount - 1 Then
            If UCase(ccKel.List(i)) Like "*" + UCase(ccKel.Text) + "*" = False Then
                ccKel.Text = ccKel.List(0)
                ccKel_Click
                Exit Sub
            End If
        End If
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
c_Ganti = 0
xID = ""
End Sub

Private Sub hTunggal_Click()
On Error Resume Next
If hTunggal.Value = 1 Then
    Frame4.Visible = True
    ccKec.Enabled = False
    ccKel.Enabled = False
    'ccKec.Text = ""
    dRekam1.Enabled = True
Else
    Frame4.Visible = False
    ccKec.Enabled = True
    ccKel.Enabled = True
    dRekam1.Enabled = False
End If
End Sub

Private Sub hTunggal_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub tNIP_GotFocus()
On Error Resume Next
tNIP.SelStart = 0
tNIP.SelLength = Len(tNIP.Text)
tNIP.SetFocus
tNIP.Alignment = 0

End Sub

Private Sub tNIP_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tNIP_LostFocus()
On Error Resume Next
If tNIP.Text = "" Or tNIP.Text = "-" Or tNIP.Text = "." Then
    tNIP.Text = 0
End If
tNIP.Alignment = 1

End Sub

Private Sub tSPPT_GotFocus()
On Error Resume Next
tSPPT.SelStart = 0
tSPPT.SelLength = Len(tSPPT.Text)
tSPPT.SetFocus
tSPPT.Alignment = 0
End Sub

Private Sub tSPPT_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tSPPT_LostFocus()
On Error Resume Next
If tSPPT.Text = "" Or tSPPT.Text = "-" Or tSPPT.Text = "." Then
    tSPPT.Text = 0
End If
tSPPT.Alignment = 1

End Sub

Private Sub tSPPT2_GotFocus()
On Error Resume Next
tSPPT2.SelStart = 0
tSPPT2.SelLength = Len(tSPPT2.Text)
tSPPT2.SetFocus
tSPPT2.Alignment = 0
End Sub

Private Sub tSPPT2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tSPPT2_LostFocus()
On Error Resume Next
If tSPPT2.Text = "" Or tSPPT2.Text = "-" Or tSPPT2.Text = "." Then
    tSPPT2.Text = 0
End If
tSPPT2.Alignment = 1
End Sub

Private Sub tTotal_GotFocus()
On Error Resume Next
tTotal.SelStart = 0
tTotal.SelLength = Len(tTotal.Text)
tTotal.SetFocus
tTotal.Alignment = 0

End Sub

Private Sub tTotal_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789-,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub tTotal_LostFocus()
On Error Resume Next
If tTotal.Text = "" Or tTotal.Text = "-" Or tTotal.Text = "." Then
    tTotal.Text = 0
End If
tTotal.Alignment = 1

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
End If
End Sub


Sub t_Normal()
On Error Resume Next
Me.Height = 4380
Me.Width = 6000
Picture1.Top = 3315
cmdOK.Top = 3480
cmdCear.Top = 3480
cmdExit.Top = 3480
Frame1.Visible = False
Label6.Visible = False
cUrut.Visible = False
cRekam.Visible = False
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
End Sub

Sub J_Karakter()
On Error GoTo Salah
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

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub



Sub cHapus_Log()
On Error Resume Next
Screen.MousePointer = vbHourglass
xxKon = "pendapatan0134"
ccKon = InputBox("Masukkan kode konfirmasi:", "Kode", "xxxx")
If UCase(Trim(xxKon)) <> UCase(Trim(ccKon)) Then MsgBox "Kode konfirmasi tidak sesuai...", vbCritical, "Error": GoTo Keluar
TANYA = MsgBox("Menghapus log data dari database...!!", vbExclamation + vbYesNo, "Logged Deleted")
If TANYA = vbNo Then GoTo Keluar
cHapus = "Delete from templogutama"
openDB (cHapus)
'ccHapus = "delete from logutama"
'openDB (ccHapus)
'KEC1 = Left(Trim(ccKec.Text), 3)
'KEL1 = Left(Trim(ccKel.Text), 3)
    If CetakQ = 1 Then
        C_STR = "DELETE  from LogUtama where NOP1='" & Trim(aNOP.Text) & "'"
    ElseIf CetakQ = 2 Then
        C_STR = "DELETE from LogUtama where (CONVERT(VARCHAR(10),TGL_PEREKAMAN_OP,110)>='" & Format(dRekam1.Value, "DD-MM-YYYY") & "' AND CONVERT(VARCHAR(10),TGL_PEREKAMAN_OP,110)<= '" & Format(dRekam2.Value, "DD-MM-YYYY") & "') "
    ElseIf CetakQ = 3 Then
        C_STR = "DELETE from LogUtama"
    ElseIf CetakQ = 4 Then
        C_STR = "DELETE from LogUtama where SUBSTRING(NOP1,7,3)='" & C_KEC & "'AND (CONVERT(VARCHAR(10),TGL_PEREKAMAN_OP,110)>='" & Format(dRekam1.Value, "DD-MM-YYYY") & "' AND CONVERT(VARCHAR(10),TGL_PEREKAMAN_OP,110)<= '" & Format(dRekam2.Value, "DD-MM-YYYY") & "')"
    ElseIf CetakQ = 5 Then
        C_STR = "DELETE from LogUtama where SUBSTRING(NOP1,7,3)='" & C_KEC & "'"
    ElseIf CetakQ = 6 Then
        C_STR = "DELETE from LogUtama where SUBSTRING(NOP1,7,3)='" & C_KEC & "' AND SUBSTRING(NOP1,11,3)='" & C_KEL & "'AND (CONVERT(VARCHAR(10),TGL_PEREKAMAN_OP,110)>='" & Format(dRekam1.Value, "DD-MM-YYYY") & "' AND CONVERT(VARCHAR(10),TGL_PEREKAMAN_OP,110)<= '" & Format(dRekam2.Value, "DD-MM-YYYY") & "')"
    ElseIf CetakQ = 7 Then
        C_STR = "DELETE from LogUtama where SUBSTRING(NOP1,7,3)='" & C_KEC & "' AND SUBSTRING(NOP1,11,3)='" & C_KEL & "'"
    End If
    openDB (C_STR)

MsgBox "Log data berhasil dihapus seluruhnya....!", vbOKOnly, "Sukses!"
Keluar:
Screen.MousePointer = vbDefault
End Sub
