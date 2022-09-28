VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKomponenBiaya2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Komponen Biaya Bangunan"
   ClientHeight    =   8490
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12195
   ControlBox      =   0   'False
   Icon            =   "frmKomponenBiaya2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   12195
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   60
      Picture         =   "frmKomponenBiaya2.frx":1CCA
      ScaleHeight     =   405
      ScaleWidth      =   12075
      TabIndex        =   25
      Top             =   15
      Width           =   12075
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
         Left            =   8580
         TabIndex        =   1
         Top             =   105
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   7065
         TabIndex        =   0
         Top             =   105
         Value           =   1  'Checked
         Width           =   1500
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   10350
         TabIndex        =   2
         Top             =   105
         Width           =   1665
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Komponen Material"
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
         Left            =   825
         TabIndex        =   26
         Top             =   105
         Width           =   1380
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
      Left            =   6360
      TabIndex        =   9
      Top             =   7935
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
      Left            =   5385
      TabIndex        =   8
      Top             =   7935
      Width           =   990
   End
   Begin MSComctlLib.ListView vBangunan 
      Height          =   6510
      Left            =   45
      TabIndex        =   15
      Top             =   915
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   11483
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
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
         Text            =   "Kode"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nama Komponen"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Satuan"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Harga Lama"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Harga Baru"
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
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "KEG"
         Object.Width           =   2540
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
      Left            =   4410
      TabIndex        =   7
      Top             =   7935
      Width           =   990
   End
   Begin VB.Frame Frame2 
      Height          =   570
      Left            =   60
      TabIndex        =   18
      Top             =   330
      Width           =   7860
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
         Index           =   1
         Left            =   1470
         TabIndex        =   3
         Top             =   180
         Width           =   1320
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
         Index           =   2
         Left            =   3645
         TabIndex        =   4
         Top             =   180
         Width           =   4155
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
         Left            =   150
         TabIndex        =   20
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Komponen"
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
         Left            =   2865
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Left            =   7935
      TabIndex        =   21
      Top             =   330
      Width           =   4170
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
         Left            =   915
         TabIndex        =   5
         Top             =   165
         Width           =   2820
      End
      Begin VB.CommandButton cmdCari 
         Height          =   360
         Left            =   3705
         Picture         =   "frmKomponenBiaya2.frx":6332
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   150
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Harga Baru"
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
         TabIndex        =   22
         Top             =   240
         Width           =   1365
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
      Left            =   2715
      TabIndex        =   12
      Top             =   8610
      Visible         =   0   'False
      Width           =   7890
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
         Left            =   1500
         TabIndex        =   16
         Top             =   525
         Width           =   6330
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
         Left            =   1500
         TabIndex        =   13
         Top             =   210
         Width           =   6330
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
         TabIndex        =   17
         Top             =   570
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
         Height          =   180
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   525
      Left            =   9525
      TabIndex        =   23
      Top             =   7335
      Width           =   2565
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
         Left            =   2115
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   150
         Width           =   345
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
         TabIndex        =   24
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Image Image1 
      Height          =   9180
      Left            =   -60
      Picture         =   "frmKomponenBiaya2.frx":6FFC
      Stretch         =   -1  'True
      Top             =   -60
      Width           =   12960
   End
End
Attribute VB_Name = "frmKomponenBiaya2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jumRek, K1, K2, PBBMin
Dim xxTahun
Private Sub cmdBangunan_Click()
frmOP_Bangunan.Show
End Sub

Private Sub cmdBumi_Click()
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
On Error GoTo Salah
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
    xTanya = MsgBox("Apa anda yakin menghapus Nilai DBKB Material?", vbQuestion + vbYesNo, "Penghapusan")
    
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
        xTanya = MsgBox("Apa anda yakin melakukan pemutakhiran Nilai DBKB Material?", vbQuestion + vbYesNo, "Pemutakhiran")
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
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

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

Private Sub cmdPersen_Click()
On Error GoTo Salah
Dim JUMP
For i = 1 To vBangunan.ListItems.Count
    JUMP = (tPersen.Text * vBangunan.ListItems.Item(i).ListSubItems(5).Text) / 100
    If vBangunan.ListItems.Item(i).ListSubItems(7).Text = "-" Then
        vBangunan.ListItems.Item(i).ListSubItems(6).Text = Format(vBangunan.ListItems.Item(i).ListSubItems(5).Text + JUMP, "#,#0.00")
    End If
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
    call_SIMPAN
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
      
        JFAS = Left(cboNOP(2).Text, 2) * 1
        vBangunan.ListItems.Clear
        STRITEM = "SELECT DBKB_MATERIAL.KD_PROPINSI, DBKB_MATERIAL.KD_DATI2, DBKB_MATERIAL.THN_DBKB_MATERIAL, DBKB_MATERIAL.KD_PEKERJAAN, PEKERJAAN.NM_PEKERJAAN, DBKB_MATERIAL.KD_KEGIATAN, PEKERJAAN_KEGIATAN.NM_KEGIATAN, DBKB_MATERIAL.NILAI_DBKB_MATERIAL FROM (PEKERJAAN INNER JOIN PEKERJAAN_KEGIATAN ON PEKERJAAN.KD_PEKERJAAN = PEKERJAAN_KEGIATAN.KD_PEKERJAAN) INNER JOIN DBKB_MATERIAL ON (PEKERJAAN.KD_PEKERJAAN = DBKB_MATERIAL.KD_PEKERJAAN) AND (PEKERJAAN_KEGIATAN.KD_KEGIATAN = DBKB_MATERIAL.KD_KEGIATAN) where DBKB_MATERIAL.KD_PEKERJAAN= '" & Left(cboNOP(2).Text, 2) & "' AND DBKB_MATERIAL.THN_DBKB_MATERIAL= '" & xxTahun & "' order by DBKB_MATERIAL.KD_PEKERJAAN asc"
        openDB (STRITEM)

        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
                vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_PEKERJAAN])
                vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", rPajak!NM_KEGIATAN
                vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", "-" 'Trim(rPajak![NM_KEGIATAN])
                vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", Format(Trim(rPajak![NILAI_DBKB_MATERIAL]), "#,#0.00")
                vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", 0
                vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
                vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", rPajak!KD_KEGIATAN
                
        'End If
        rPajak.MoveNext
        Loop
        tBumi(0).Text = 0
    
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
'strK1 = "Select * From GROUP_RESOURCE order by KD_GROUP_RESOURCE ASC"
strK1 = "Select * From PEKERJAAN ORDER BY KD_PEKERJAAN ASC "
openDB (strK1)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    'cboNOP(2).AddItem rPajak!KD_GROUP_RESOURCE & "-" & rPajak!NM_GROUP_RESOURCE
    If Trim(rPajak!KD_PEKERJAAN) = "21" Or Trim(rPajak!KD_PEKERJAAN) = "22" Or Trim(rPajak!KD_PEKERJAAN) = "23" Or Trim(rPajak!KD_PEKERJAAN) = "24" Then
        cboNOP(2).AddItem rPajak!KD_PEKERJAAN & " " & rPajak!NM_PEKERJAAN
    End If
    
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

End Sub


Private Sub tBumi_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
Select Case Index
    Case 0
        If KeyAscii = 13 Then
            cmdCari_Click
        End If
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

Private Sub vBangunan_Click()
On Error GoTo Salah
'If vBangunan.SelectedItem.ListSubItems(5).Text = 1 Then
    For i = 1 To vBangunan.ListItems.Count
        vBangunan.ListItems.Item(i).ListSubItems(8).Text = "-"
    Next
 vBangunan.SelectedItem.ListSubItems(8).Text = "Proses"
    tBumi(0).Text = vBangunan.SelectedItem.ListSubItems(5).Text
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
                'tBumi(0).SetFocus
            End If
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Description, vbCritical, "TETNONG: " & Err.Number
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

Private Sub vBangunan_KeyDown(KeyCode As Integer, Shift As Integer)
vBangunan_Click
End Sub

Private Sub vBangunan_KeyUp(KeyCode As Integer, Shift As Integer)
vBangunan_Click
End Sub
Private Sub vBangunan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan.SortKey = ColumnHeader.Index - 1
vBangunan.Sorted = True
vBangunan.Sorted = False
vBangunan.SortOrder = lvwAscending
End Sub

Sub bersih()
tBumi(0).Text = 0
vBangunan.ListItems.Clear
End Sub

Sub call_SIMPAN()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM DBKB_MATERIAL WHERE THN_DBKB_MATERIAL ='" & Trim(cboNOP(1).Text) & "' AND KD_PEKERJAAN='" & Left(Trim(cboNOP(2).Text), 2) & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Daftar Biaya Komponen Bangunan (DBKB) Material: " & cboNOP(1).Text & _
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
If (cmdSave.Caption = "&Update" And chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" And chPajak(2).Value = 1) Then

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
        rPajak!THN_DBKB_MATERIAL = cboNOP(1).Text
        rPajak!KD_PEKERJAAN = Left(Trim(cboNOP(2).Text), 2)
        rPajak!KD_KEGIATAN = vBangunan.ListItems.Item(i).ListSubItems(9).Text
        rPajak!NILAI_DBKB_MATERIAL = vBangunan.ListItems.Item(i).ListSubItems(6).Text

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


