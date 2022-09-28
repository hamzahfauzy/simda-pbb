VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPejabat 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA PEJABAT BERWENANG"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6150
   Begin VB.TextBox txID 
      Height          =   315
      Left            =   4260
      TabIndex        =   17
      Top             =   4170
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
      Left            =   4140
      TabIndex        =   7
      Top             =   4440
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
      Left            =   3300
      TabIndex        =   6
      Top             =   4440
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
      Left            =   2460
      TabIndex        =   5
      Top             =   4440
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
      Left            =   1620
      TabIndex        =   3
      Top             =   4440
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
      Left            =   780
      TabIndex        =   4
      Top             =   4440
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
      Height          =   1305
      Left            =   -15
      TabIndex        =   14
      Top             =   -90
      Width           =   6165
      Begin VB.TextBox tJabatan 
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
         Top             =   915
         Width           =   4440
      End
      Begin VB.TextBox tNIP 
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
         TabIndex        =   0
         Top             =   225
         Width           =   3150
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
         Top             =   570
         Width           =   4440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
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
         TabIndex        =   18
         Top             =   930
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Top             =   570
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIP"
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
         TabIndex        =   15
         Top             =   240
         Width           =   285
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
      Height          =   3075
      Left            =   -15
      TabIndex        =   11
      Top             =   1110
      Width           =   6165
      Begin MSComctlLib.ListView vAnak 
         Height          =   2775
         Left            =   105
         TabIndex        =   8
         Top             =   195
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   4895
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "a"
            Text            =   "No"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "b"
            Text            =   "NIP"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "NAMA"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "JABATAN"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   900
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin VB.Label LPas 
      Caption         =   "Label1"
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   3075
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LLevel 
      Caption         =   "Label3"
      Height          =   285
      Left            =   270
      TabIndex        =   9
      Top             =   3510
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "frmPejabat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cEdit



Private Sub cmdBaru_Click()
On Error Resume Next
Baru
vAnak.ListItems.Clear
TampilUser
cmdHapus.Enabled = False
cEdit = 0
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
cEdit = 1
'Edha.Visible = True
cmdEdit.Enabled = False
cmdHapus.Enabled = True
cmdSimpan.Caption = "&Update"
cmdBaru.Caption = "&Batal"
'Edha.SetFocus
vAnak.Enabled = True
End Sub

Private Sub cmdHapus_Click()
On Error Resume Next
Hapus
TampilUser
Lama
End Sub

Private Sub cmdKeluar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
For Each Control In Me
    If TypeOf Control Is TextBox Then
        If Control.Text = "" Then
            MsgBox "Data tidak boleh kosong..." & Control.Name, vbCritical, "Error"
            tUser.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
Next
C_STR = "SELECT * FROM PEJABAT WHERE NIP+'-'+NAMA='" & Trim(txID.Text) & "'"
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
        E_STR = "UPDATE [PEJABAT] SET [NAMA]='" & tNama.Text & "',[NIP]='" & tNIP.Text & "' ,[JABATAN]='" & tJabatan.Text & "' WHERE  ([NIP]+'-'+NAMA)='" & txID.Text & "' "
        openDB (E_STR)
    Else
        I_STR = "INSERT INTO [PEJABAT] ([NIP],[NAMA],[JABATAN]) VALUES ('" & tNIP.Text & "','" & tNama.Text & "','" & tJabatan.Text & "' )"
        openDB (I_STR)
    End If
TampilUser
Lama
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub cSKPD_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub



Private Sub Form_Activate()
On Error Resume Next
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
Me.Left = (frmUtama.Width - Me.Width) / 2
Lama
vAnak.Enabled = False
cmdHapus.Enabled = False
End Sub

Sub Baru()
On Error Resume Next
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
tNIP.SetFocus
'Edha.Visible = False
End Sub
Sub Lama()
On Error Resume Next
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
 If InStr("123456789.-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
End Sub

Private Sub tLevel_LostFocus()
On Error Resume Next
If tLevel.Text = 1 Then
    cSKPD.Text = "Administrator"
End If
End Sub

Private Sub tJabatan_GotFocus()
On Error Resume Next
tJabatan.SelStart = 0
tJabatan.SelLength = Len(tJabatan.Text)
End Sub

Private Sub tJabatan_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    SendKeys "{tab}"
     KeyAscii = 0
End If

End Sub

Private Sub tJabatan_LostFocus()
tJabatan.Text = Rep(tJabatan.Text)
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

Private Sub tNama_LostFocus()
tNama.Text = Rep(tNama.Text)
End Sub

Private Sub tNIP_GotFocus()
On Error Resume Next
tNIP.SelStart = 0
tNIP.SelLength = Len(tNIP.Text)
End Sub

Private Sub tNIP_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
     SendKeys "{tab}"
     KeyAscii = 0
     
End If
If InStr("0123456789 -,.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub

Private Sub vAnak_Click()
On Error Resume Next
  tNIP.Text = vAnak.SelectedItem.ListSubItems(1).Text
  tNama.Text = vAnak.SelectedItem.ListSubItems(2).Text
  tJabatan.Text = vAnak.SelectedItem.ListSubItems(3).Text
  txID.Text = vAnak.SelectedItem.ListSubItems(4).Text
  
End Sub
Sub TampilUser()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
vAnak.ListItems.Clear
C_STR = "select * from [PEJABAT] ORDER BY NIP ASC"
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'rsAdmin.Find "NIP= '" & tNIP.Text & "'"
i = 0
Do While Not rPajak.EOF
    i = i + 1
    vAnak.ListItems.Add i, "", Format(i, "000")
    vAnak.ListItems.Item(i).ListSubItems.Add 1, "", rPajak![NIP]
    vAnak.ListItems.Item(i).ListSubItems.Add 2, "", rPajak![NAMA]
    vAnak.ListItems.Item(i).ListSubItems.Add 3, "", rPajak![JABATAN]
    vAnak.ListItems.Item(i).ListSubItems.Add 4, "", rPajak![NIP] & "-" & rPajak![NAMA]
    rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub


Sub Hapus()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
C_SSTR = "select * from PEJABAT"
openDB (C_SSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If rPajak.RecordCount = 1 Then MsgBox "Pejabat tidak dapat dihapus...!", vbCritical, "Error": Screen.MousePointer = vbDefault: Exit Sub
C_STR = "DELETE From [PEJABAT] where [NIP]+'-'+[NAMA]='" & txID.Text & "'"
openDB (C_STR)
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub vAnak_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vAnak.SortKey = ColumnHeader.Index - 1
vAnak.Sorted = True
vAnak.SortOrder = lvwAscending
vAnak.Sorted = False
End Sub

Private Sub vAnak_KeyDown(KeyCode As Integer, Shift As Integer)
vAnak_Click
End Sub

Private Sub vAnak_KeyUp(KeyCode As Integer, Shift As Integer)
vAnak_Click
End Sub
