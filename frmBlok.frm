VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBlok 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Blok"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6645
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
      TabIndex        =   14
      Top             =   -105
      Width           =   6645
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
         Left            =   4650
         TabIndex        =   2
         Top             =   270
         Visible         =   0   'False
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
         Left            =   2265
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
      Left            =   3705
      TabIndex        =   8
      Top             =   5565
      Width           =   915
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
      Left            =   2805
      TabIndex        =   7
      Top             =   5565
      Width           =   915
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
      Left            =   1905
      TabIndex        =   6
      Top             =   5565
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   1350
      Left            =   0
      TabIndex        =   10
      Top             =   390
      Width           =   6645
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
         ItemData        =   "frmBlok.frx":0000
         Left            =   1845
         List            =   "frmBlok.frx":0002
         TabIndex        =   4
         Top             =   570
         Width           =   4575
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
         ItemData        =   "frmBlok.frx":0004
         Left            =   1845
         List            =   "frmBlok.frx":0006
         TabIndex        =   3
         Top             =   210
         Width           =   4575
      End
      Begin VB.TextBox tBlok 
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
         TabIndex        =   5
         Top             =   945
         Width           =   1965
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Blok"
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
         TabIndex        =   13
         Top             =   990
         Width           =   1320
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   615
         Width           =   1320
      End
      Begin VB.Label Label1 
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
         Left            =   150
         TabIndex        =   11
         Top             =   270
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView vBangunan 
      Height          =   3765
      Left            =   0
      TabIndex        =   9
      Top             =   1725
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   6641
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
      NumItems        =   6
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
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "KD_KEC"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "KD_KEL"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "BLOK"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "KELURAHAN"
         Object.Width           =   4410
      EndProperty
   End
End
Attribute VB_Name = "frmBlok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ccKec_Click()
On Error Resume Next
CALL_KEL
End Sub

Private Sub ccKec_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
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

Private Sub ccKel_Click()
On Error Resume Next
CALL_BLOK
End Sub

Private Sub ccKel_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
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

Private Sub chPajak_Click(Index As Integer)
On Error GoTo Salah
If chPajak(1).Value = 0 Then
        If chPajak(2).Value = 0 And chPajak(3).Value = 0 Then chPajak(1).Value = 1
        
    End If
Select Case Index
Case 1
    If chPajak(1).Value = 1 Then
        chPajak(2).Value = 0: chPajak(3).Value = 0
        cmdSave.Caption = "&Save"
        cmdClear_Click
    End If
Case 2
    If chPajak(2).Value = 1 Then
        chPajak(1).Value = 0: chPajak(3).Value = 0
        cmdSave.Caption = "&Delete"
    End If
'tKode.Locked = True
Case 3
    If chPajak(3).Value = 1 Then
        chPajak(1).Value = 0: chPajak(2).Value = 0
        cmdSave.Caption = "&Update"
    End If
'tKode.Locked = True
End Select
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
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
ccKec.Text = ""
ccKel.Text = ""
tBlok.Text = ""
ccKec.SetFocus
vBangunan.ListItems.Clear
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo Salah
Dim CD_KEC, CD_KEL
CD_KEC = Left(Trim(ccKec.Text), 3)
CD_KEL = Left(Trim(ccKel.Text), 3)
B_SQL = "SELECT KD_KECAMATAN,KD_KELURAHAN,KD_BLOK FROM DAT_OP_BUMI WHERE KD_KECAMATAN='" & CD_KEC & "'  AND KD_KELURAHAN='" & CD_KEL & "'  AND KD_BLOK='" & tBlok.Text & "'"
openDB (B_SQL)
If rPajak.RecordCount > 0 Then rPajak.MoveNext
If Not rPajak.EOF And chPajak(2).Value = 1 Then
    MsgBox "ANDA TIDAK DAPAT MENGHAPUS BLOK INI", vbCritical, "TETNONG"
    Exit Sub
End If
If ccKec.Text = "" Or ccKel.Text = "" Or tBlok.Text = "" Then
     MsgBox "MASIH ADA DATA YANG KOSONG...", vbCritical, "TETNONG"
       ' ccKec.SetFocus
            Exit Sub
End If
iSQL = "Select * From DAT_PETA_BLOK WHERE KD_KECAMATAN='" & CD_KEC & "'  AND KD_KELURAHAN='" & CD_KEL & "'  AND KD_BLOK='" & tBlok.Text & "'  ORDER BY KD_BLOK,KD_KELURAHAN, KD_KECAMATAN ASC"
openDB (iSQL)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If chPajak(1).Value = 1 Then
xTanya = MsgBox("APA ANDA YAKIN MENYIMPAN DATA INI?", vbInformation + vbYesNo, "Info")
If xTanya = vbNo Then Exit Sub
    If Not rPajak.EOF Then
        MsgBox "Data Sudah Ada, Silahkan Diganti...", vbCritical, "Tetnong"
        Exit Sub
    End If
    iSQL = "Insert Into DAT_PETA_BLOK VALUES ('12','12','" & CD_KEC & "','" & CD_KEL & "','" & tBlok.Text & "','0')"
    openDB (iSQL)
ElseIf chPajak(2).Value = 1 Then
xTanya = MsgBox("APA ANDA YAKIN MENGHAPUS DATA INI?", vbInformation + vbYesNo, "Info")
If xTanya = vbNo Then Exit Sub

    If rPajak.EOF Then
        MsgBox "Data Tidak Ada...", vbCritical, "Tetnong"
        Exit Sub
    End If
    iSQL = "Delete From DAT_PETA_BLOK WHERE KD_KECAMATAN='" & CD_KEC & "'  AND KD_KELURAHAN='" & CD_KEL & "'  AND KD_BLOK='" & tBlok.Text & "'"
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
    rPajak!KD_KECAMATAN = CD_KEC
    rPajak!KD_KELURAHAN = CD_KEL
    rPajak!KD_BLOK = tBlok.Text
    rPajak.Update
    'iSQL = "UPDATE REF_KECAMATAN SET KD_PROPINSI='12',KD_DATI2='12',KD_KECAMATAN='" & tKode.Text & "',NM_KECAMATAN='" & tKec.Text & "' where KD_KECAMATAN='" & tKode.Text & "'"
    'openDB (iSQL)
End If
CALL_BLOK
ccKec.Text = ""
ccKel.Text = ""
tBlok.Text = ""
'ccKec.SetFocus
'CALL_BLOK
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
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
ccKec.Clear
QSTR = "SELECT * FROM REF_KECAMATAN ORDER BY KD_KECAMATAN ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            ccKec.AddItem rPajak!KD_KECAMATAN & " " & rPajak!NM_KECAMATAN
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_KEL()
On Error GoTo Salah
ccKel.Clear
QSTR = "SELECT * FROM REF_KELURAHAN WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "' ORDER BY KD_KELURAHAN ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            ccKel.AddItem rPajak!KD_KELURAHAN & " " & rPajak!NM_KELURAHAN
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_BLOK()
On Error GoTo Salah
vBangunan.ListItems.Clear
QSTR = "SELECT * FROM DAT_PETA_BLOK WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' ORDER BY KD_KELURAHAN ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_KECAMATAN])
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Trim(rPajak![KD_KELURAHAN])
        'vBangunan.ListItems.Item(I).ListSubItems.Add 4, "", Mid(ccKel.Text, 5, Len(ccKel.Text))
        vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", Trim(rPajak![KD_BLOK])
        
                
        rPajak.MoveNext
        Loop
QSTR1 = "SELECT * FROM REF_KELURAHAN WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "' ORDER BY KD_KELURAHAN ASC"
openDB (QSTR1)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            For i = 1 To vBangunan.ListItems.Count
                If rPajak!KD_KELURAHAN = vBangunan.ListItems.Item(i).ListSubItems(3).Text Then
                    vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", rPajak!NM_KELURAHAN
                End If
            Next
        rPajak.MoveNext
        Loop

If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If

End Sub

Private Sub tBlok_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub vBangunan_Click()
On Error GoTo Salah
If chPajak(2).Value = 1 Or chPajak(3).Value = 1 Then
    tBlok.Text = vBangunan.SelectedItem.ListSubItems(4).Text
    'If vBangunan.SelectedItem.ListSubItems(3).Text = Left(Trim(ccKel.Text), 3) Then
    For i = 0 To ccKec.ListCount - 1
        'ccKec.Text = vBangunan.SelectedItem.ListSubItems(2).Text
        If vBangunan.SelectedItem.ListSubItems(2).Text = Left(Trim(ccKec.List(i)), 3) Then
           ccKec.Text = ccKec.List(i)
        End If
    Next
    ccKel.Text = vBangunan.SelectedItem.ListSubItems(3).Text & " " & vBangunan.SelectedItem.ListSubItems(5).Text
End If

If Err.Number = 0 Then Exit Sub
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub vBangunan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan.SortKey = ColumnHeader.Index - 1
vBangunan.Sorted = True
vBangunan.Sorted = False
vBangunan.SortOrder = lvwAscending
End Sub

Private Sub vBangunan_KeyDown(KeyCode As Integer, Shift As Integer)
vBangunan_Click
End Sub

Private Sub vBangunan_KeyUp(KeyCode As Integer, Shift As Integer)
vBangunan_Click
End Sub
