VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKec 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Kecamatan"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6285
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
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
      TabIndex        =   12
      Top             =   -105
      Width           =   6285
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
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   1
         Top             =   270
         Width           =   1695
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
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   0
         Top             =   270
         Value           =   1  'Checked
         Width           =   1935
      End
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
         Height          =   210
         Index           =   3
         Left            =   4425
         TabIndex        =   2
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
      Height          =   435
      Left            =   3510
      TabIndex        =   7
      Top             =   4665
      Width           =   990
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
      Height          =   435
      Left            =   2535
      TabIndex        =   6
      Top             =   4665
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
      Left            =   1560
      TabIndex        =   5
      Top             =   4665
      Width           =   990
   End
   Begin VB.Frame Frame2 
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
      Height          =   1005
      Left            =   0
      TabIndex        =   8
      Top             =   390
      Width           =   6285
      Begin VB.TextBox tKode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         MaxLength       =   3
         TabIndex        =   3
         Top             =   210
         Width           =   1905
      End
      Begin VB.TextBox tKec 
         Appearance      =   0  'Flat
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
         Top             =   570
         Width           =   4350
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
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
         Left            =   150
         TabIndex        =   10
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   630
         Width           =   1320
      End
   End
   Begin MSComctlLib.ListView vBangunan 
      Height          =   3045
      Left            =   0
      TabIndex        =   11
      Top             =   1380
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   5371
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Kode"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NAMA KECAMATAN"
         Object.Width           =   10583
      EndProperty
   End
End
Attribute VB_Name = "frmKec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chPajak_Click(Index As Integer)
On Error GoTo Salah
cmdClear_Click
If chPajak(1).Value = 0 Then
        If chPajak(2).Value = 0 And chPajak(3).Value = 0 Then chPajak(1).Value = 1
        
    End If
Select Case Index
Case 1
    If chPajak(1).Value = 1 Then
        chPajak(2).Value = 0: chPajak(3).Value = 0
        cmdSave.Caption = "&Save"
    End If
Case 2
    If chPajak(2).Value = 1 Then
        chPajak(1).Value = 0: chPajak(3).Value = 0
        cmdSave.Caption = "&Delete"
    End If
tKode.Locked = True
Case 3
    If chPajak(3).Value = 1 Then
        chPajak(1).Value = 0: chPajak(2).Value = 0
        cmdSave.Caption = "&Update"
    End If
tKode.Locked = True
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

Private Sub cmdClear_Click()
On Error Resume Next
tKode.Text = ""
tKec.Text = ""
vBangunan.Refresh
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo Salah
B_SQL = "SELECT KD_KECAMATAN FROM DAT_OP_BUMI WHERE KD_KECAMATAN='" & Trim(tKode.Text) & "' ORDER BY KD_KECAMATAN ASC"
openDB (B_SQL)
If rPajak.RecordCount > 0 Then rPajak.MoveNext
If Not rPajak.EOF And (chPajak(2).Value = 1 Or chPajak(3).Value = 1) Then
    MsgBox "ANDA TIDAK DAPAT MENGHAPUS/EDIT KECAMATAN INI, " & _
        vbCrLf & "KARENA SUDAH ADA TRANSAKSI...", vbCritical, "TETNONG"
        Exit Sub
End If
If tKode.Text = "" Or tKec.Text = "" Then
     MsgBox "MASIH ADA DATA YANG KOSONG...", vbCritical, "TETNONG"
        tKode.SetFocus
            Exit Sub
End If
iSQL = "Select * From REF_KECAMATAN where KD_KECAMATAN='" & tKode.Text & "'"
openDB (iSQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If chPajak(1).Value = 1 Then
xTanya = MsgBox("APA ANDA YAKIN MENYIMPAN DATA INI?", vbInformation + vbYesNo, "Info")
If xTanya = vbNo Then Exit Sub
    If Not rPajak.EOF Then
        MsgBox "Data Sudah Ada, Silahkan Diganti...", vbCritical, "Tetnong"
        Exit Sub
    End If
    iSQL = "Insert Into REF_KECAMATAN VALUES ('12','12','" & tKode.Text & "','" & tKec.Text & "')"
    openDB (iSQL)
ElseIf chPajak(2).Value = 1 Then
xTanya = MsgBox("APA ANDA YAKIN MENGHAPUS DATA INI?", vbInformation + vbYesNo, "Info")
If xTanya = vbNo Then Exit Sub

    If rPajak.EOF Then
        MsgBox "Data Tidak Ada...", vbCritical, "Tetnong"
        Exit Sub
    End If
    iSQL = "Delete From REF_KECAMATAN where KD_KECAMATAN='" & tKode.Text & "'"
    openDB (iSQL)
ElseIf chPajak(3).Value = 1 Then
xTanya = MsgBox("APA ANDA YAKIN MENGEDIT DATA INI?", vbInformation + vbYesNo, "Info")
If xTanya = vbNo Then Exit Sub

    If rPajak.EOF Then
        MsgBox "Data Tidak Ada...", vbCritical, "Tetnong"
        Exit Sub
    End If
    rPajak!KD_PROPINSI = "12"
    rPajak!KD_DATI2 = "12"
    rPajak!KD_KECAMATAN = tKode.Text
    rPajak!NM_KECAMATAN = tKec.Text
    rPajak.Update
    'iSQL = "UPDATE REF_KECAMATAN SET KD_PROPINSI='12',KD_DATI2='12',KD_KECAMATAN='" & tKode.Text & "',NM_KECAMATAN='" & tKec.Text & "' where KD_KECAMATAN='" & tKode.Text & "'"
    'openDB (iSQL)
End If
vBangunan.ListItems.Clear
CALL_KEC
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
CALL_KEC
End Sub

Sub CALL_KEC()
On Error GoTo Salah
QSTR = "SELECT * FROM REF_KECAMATAN ORDER BY KD_KECAMATAN ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_KECAMATAN])
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", rPajak![NM_KECAMATAN]
                
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub tKec_GotFocus()
On Error Resume Next
tKec.SelStart = 0
tKec.SelLength = Len(tKec.Text)

End Sub

Private Sub tKec_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub tKec_LostFocus()
tKec.Text = Rep(tKec.Text)
End Sub

Private Sub tKode_GotFocus()
On Error Resume Next
tKode.SelStart = 0
tKode.SelLength = Len(tKode.Text)

tKode.Alignment = 0
End Sub

Private Sub tKode_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
End Sub

Private Sub tKode_LostFocus()
On Error Resume Next
tKode.Alignment = 1
End Sub

Private Sub vBangunan_Click()
On Error Resume Next
If chPajak(2).Value = 1 Or chPajak(3).Value = 1 Then
    tKode.Text = vBangunan.SelectedItem.ListSubItems(2).Text
    tKec.Text = vBangunan.SelectedItem.ListSubItems(3).Text
Else
    cmdClear_Click
End If
End Sub

Private Sub vBangunan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan.SortKey = ColumnHeader.Index - 1
vBangunan.Sorted = True
vBangunan.Sorted = False
vBangunan.SortOrder = lvwAscending
End Sub

Sub CALL_BUMI()
End Sub
Private Sub vBangunan_KeyDown(KeyCode As Integer, Shift As Integer)
vBangunan_Click
End Sub

Private Sub vBangunan_KeyUp(KeyCode As Integer, Shift As Integer)
vBangunan_Click
End Sub

