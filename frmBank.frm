VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBank 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tempat Pembayaran PBB"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6150
   Begin VB.TextBox txID 
      Height          =   315
      Left            =   4260
      TabIndex        =   18
      Top             =   4875
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "&Keluar"
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
      Left            =   4245
      TabIndex        =   8
      Top             =   5745
      Width           =   855
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
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
      Left            =   3405
      TabIndex        =   7
      Top             =   5745
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
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
      Left            =   2565
      TabIndex        =   6
      Top             =   5745
      Width           =   855
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
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
      Left            =   1725
      TabIndex        =   4
      Top             =   5745
      Width           =   855
   End
   Begin VB.CommandButton cmdBaru 
      Caption         =   "&Baru"
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
      Left            =   885
      TabIndex        =   5
      Top             =   5745
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   -15
      TabIndex        =   15
      Top             =   -90
      Width           =   6165
      Begin VB.TextBox tRek 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1575
         TabIndex        =   3
         Top             =   1230
         Width           =   4440
      End
      Begin VB.TextBox tAlamat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1575
         TabIndex        =   2
         Top             =   870
         Width           =   4440
      End
      Begin VB.TextBox tKode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   2
         TabIndex        =   0
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox tNama 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1575
         TabIndex        =   1
         Top             =   525
         Width           =   4440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Rekening"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   20
         Top             =   1275
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   19
         Top             =   915
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama TP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   17
         Top             =   570
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode TP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   16
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4035
      Left            =   -15
      TabIndex        =   12
      Top             =   1545
      Width           =   6165
      Begin MSComctlLib.ListView vAnak 
         Height          =   3720
         Left            =   75
         TabIndex        =   9
         Top             =   180
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   6562
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483642
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "a"
            Text            =   "No"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Key             =   "b"
            Text            =   "KD_TP"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "NM_TP"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ALAMAT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "NO_REK"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Pangkat/Gol/Ruang"
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
         Left            =   345
         TabIndex        =   14
         Top             =   930
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Kerja"
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
         Left            =   120
         TabIndex        =   13
         Top             =   900
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin VB.Label LPas 
      Caption         =   "Label1"
      Height          =   300
      Left            =   240
      TabIndex        =   11
      Top             =   3075
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LLevel 
      Caption         =   "Label3"
      Height          =   285
      Left            =   270
      TabIndex        =   10
      Top             =   3510
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cEdit



Private Sub cmdBaru_Click()
On Error GoTo Salah
Baru
vAnak.ListItems.Clear
TampilUser
cmdHapus.Enabled = False
cEdit = 0
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

End Sub

Private Sub cmdEdit_Click()
On Error GoTo Salah
cEdit = 1
'Edha.Visible = True
cmdEdit.Enabled = False
cmdHapus.Enabled = True
cmdSimpan.Caption = "&Update"
cmdBaru.Caption = "&Batal"
'Edha.SetFocus
vAnak.Enabled = True
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

End Sub

Private Sub cmdHapus_Click()
On Error GoTo Salah
Hapus
TampilUser
'Lama
tKode.Text = "": tNama.Text = "": tAlamat.Text = "": tRek.Text = ""
vAnak.SetFocus
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

End Sub

Private Sub cmdKeluar_Click()

Unload Me
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
For Each Control In Me
    If TypeOf Control Is TextBox Then
        If Control.Text = "" Then
            MsgBox "Data tidak boleh kosong..." & Control.Name, vbCritical, "Error"
            tNama.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
Next
C_STR = "SELECT * FROM TEMPAT_BAYAR WHERE NM_TP +'-'+ ALAMAT_TP +'-'+ KD_TP ='" & Trim(txID.Text) & "'"
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then
    If cEdit = 0 Then
        MsgBox "DATA SUDAH ADA...!", vbCritical, "TETNONG...!"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
Else
    If cEdit = 1 Then
        MsgBox "DATA TIDAK ADA...!", vbCritical, "TETNONG...!"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End If
    If cEdit = 1 Then
        'E_STR = "UPDATE [TEMPAT_BAYAR] SET KD_KANWIL='01',KD_KPPBB='16',KD_BANK_TUNGGAL='01',KD_BANK_PERSEPSI='01',[KD_TP]='" & tKode.Text & "',[NM_TP]='" & tNama.Text & "' ,[ALAMAT_TP]='" & tAlamat.Text & "' ,[NO_REK_TP]='" & tRek.Text & "' WHERE  ([NM_TP]+'-'+ALAMAT_TP+'-'+KD_TP)='" & txID.Text & "' "
        E_STR = "SP_UPDATE_BANK '01','16','01','01','" & tKode.Text & "','" & tNama.Text & "','" & tAlamat.Text & "','" & tRek.Text & "','" & txID.Text & "' "
        openDB (E_STR)
    Else
        'I_STR = "INSERT INTO [TEMPAT_BAYAR] (KD_KANWIL,KD_KPPBB,KD_BANK_TUNGGAL,KD_BANK_PERSEPSI,[KD_TP],[NM_TP],ALAMAT_TP,NO_REK_TP) VALUES ('01','16','01','01','" & tKode.Text & "','" & tNama.Text & "','" & tAlamat.Text & "','" & tRek.Text & "')"
        I_STR = "SP_INSERT_BANK '01','16','01','01','" & tKode.Text & "','" & tNama.Text & "','" & tAlamat.Text & "','" & tRek.Text & "'"
            openDB (I_STR)
    End If
TampilUser
Lama
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:
Screen.MousePointer = vbDefault
End Sub

Private Sub cSKPD_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub



Private Sub Form_Activate()
On Error GoTo Salah
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
Lama
vAnak.Enabled = False
cmdHapus.Enabled = False
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

End Sub

Sub Baru()
On Error GoTo Salah
For Each Control In Me
    If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
        Control.BackColor = vbWhite
        Control.Enabled = True
        Control.Locked = False
        Control.Text = ""
    End If
Next
txID.Text = "-"
cmdSimpan.Enabled = True
cmdBaru.Enabled = True
cmdEdit.Enabled = True
cmdHapus.Enabled = True
cmdSimpan.Caption = "&Simpan"
cmdBaru.Caption = "&Baru"
tKode.SetFocus
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

'Edha.Visible = False
End Sub
Sub Lama()
On Error GoTo Salah
For Each Control In Me
    If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
        Control.BackColor = vbButtonFace
        Control.Enabled = False
        
    End If
Next
cmdSimpan.Enabled = False
cmdBaru.Enabled = True
cmdEdit.Enabled = False
cmdHapus.Enabled = False
cmdSimpan.Caption = "&Simpan"
cmdBaru.Caption = "&Baru"
'Edha.Visible = False
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

End Sub




Private Sub tLevel_GotFocus()
On Error Resume Next
SendKeys "{Home}+{end}"
End Sub

Private Sub tLevel_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
 If InStr("123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
End Sub

Private Sub tLevel_LostFocus()
On Error Resume Next
If tLevel.Text = 1 Then
    cSKPD.Text = "Administrator"
End If
End Sub

Private Sub tAlamat_GotFocus()
On Error Resume Next
tAlamat.SelStart = 0
tAlamat.SelLength = Len(tAlamat.Text)
End Sub

Private Sub tAlamat_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub tAlamat_LostFocus()
tAlamat.Text = Rep(tAlamat.Text)
End Sub

Private Sub tKode_GotFocus()
On Error Resume Next
tKode.SelStart = 0
tKode.SelLength = Len(tKode.Text)
End Sub

Private Sub tKode_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
     KeyAscii = 0
End If
If InStr("0123456789.,", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub



Private Sub tPas_GotFocus()
On Error Resume Next
SendKeys "{Home}+{end}"
End Sub

Private Sub tPas_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub trPas_GotFocus()
On Error Resume Next
SendKeys "{Home}+{end}"
End Sub

Private Sub trPas_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If

End Sub

Private Sub tUser_GotFocus()
On Error Resume Next
SendKeys "{Home}+{end}"
End Sub

Private Sub tUser_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub tNama_GotFocus()
On Error Resume Next
tNama.SelStart = 0
tNama.SelLength = Len(tNama.Text)
End Sub

Private Sub tNama_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
     SendKeys "{tab}"
     KeyAscii = 0
     
End If
End Sub

Private Sub tNama_LostFocus()
tNama.Text = Rep(tNama.Text)
End Sub

Private Sub tRek_GotFocus()
On Error Resume Next
tRek.SelStart = 0
tRek.SelLength = Len(tRek.Text)
End Sub

Private Sub tRek_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub tRek_LostFocus()
tRek.Text = Rep(tRek.Text)
End Sub

Private Sub vAnak_Click()
On Error GoTo Salah
  tKode.Text = vAnak.SelectedItem.ListSubItems(1).Text
  tNama = vAnak.SelectedItem.ListSubItems(2).Text
  tAlamat = vAnak.SelectedItem.ListSubItems(3).Text
  tRek = vAnak.SelectedItem.ListSubItems(4).Text
  txID.Text = vAnak.SelectedItem.ListSubItems(5).Text
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

End Sub
Sub TampilUser()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
vAnak.ListItems.Clear
C_STR = "select * from [TEMPAT_BAYAR] ORDER BY KD_TP ASC"
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'rsAdmin.Find "NIP= '" & tNama.Text & "'"
i = 0
Do While Not rPajak.EOF
    i = i + 1
    vAnak.ListItems.Add i, "", Format(i, "000")
    If IsNull(rPajak!KD_TP) = True Or rPajak!KD_TP = "" Then rPajak!KD_TP = "00"
    vAnak.ListItems.Item(i).ListSubItems.Add 1, "", rPajak![KD_TP]
    If IsNull(rPajak!NM_TP) = True Or rPajak!NM_TP = "" Then rPajak!NM_TP = "-"
    vAnak.ListItems.Item(i).ListSubItems.Add 2, "", rPajak![NM_TP]
    If IsNull(rPajak!ALAMAT_TP) = True Or rPajak!ALAMAT_TP = "" Then rPajak!ALAMAT_TP = "-"
    vAnak.ListItems.Item(i).ListSubItems.Add 3, "", rPajak![ALAMAT_TP]
    If IsNull(rPajak!NO_REK_TP) = True Or rPajak!NO_REK_TP = "" Then rPajak!NO_REK_TP = "0000"
    vAnak.ListItems.Item(i).ListSubItems.Add 4, "", rPajak![NO_REK_TP]
    vAnak.ListItems.Item(i).ListSubItems.Add 5, "", rPajak![NM_TP] & "-" & rPajak![ALAMAT_TP] & "-" & rPajak![KD_TP]
    rPajak.MoveNext
Loop
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

Screen.MousePointer = vbDefault
End Sub


Sub Hapus()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
'C_STR = "DELETE From [TEMPAT_BAYAR] where [NM_TP]+'-'+ [ALAMAT_TP] +'-'+[KD_TP]='" & txID.Text & "'"
C_STR = "SP_DEL_BANK '" & txID.Text & "'"
openDB (C_STR)
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:
Screen.MousePointer = vbDefault
End Sub

Private Sub vAnak_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo Salah
vAnak.SortKey = ColumnHeader.Index - 1
vAnak.Sorted = True
vAnak.SortOrder = lvwAscending
vAnak.Sorted = False
If Err.Number = 0 Then GoTo Keluar
Salah:
If Err.Number = 0 Then GoTo Keluar
MsgBox Err.Number & ": " & Err.Description
Keluar:

End Sub

Private Sub vAnak_KeyDown(KeyCode As Integer, Shift As Integer)
vAnak_Click
End Sub

Private Sub vAnak_KeyUp(KeyCode As Integer, Shift As Integer)
vAnak_Click
End Sub
