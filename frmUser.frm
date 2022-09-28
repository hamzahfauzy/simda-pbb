VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT/DELETE/INSERT NEW USERS"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   6150
   Begin VB.TextBox txID 
      Height          =   315
      Left            =   4245
      TabIndex        =   26
      Top             =   2265
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
      Left            =   4290
      TabIndex        =   11
      Top             =   6810
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
      Left            =   3450
      TabIndex        =   10
      Top             =   6810
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
      Left            =   2610
      TabIndex        =   9
      Top             =   6810
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
      Left            =   1770
      TabIndex        =   7
      Top             =   6810
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
      Left            =   930
      TabIndex        =   8
      Top             =   6810
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
      Height          =   1335
      Left            =   -15
      TabIndex        =   22
      Top             =   -90
      Width           =   6165
      Begin VB.ComboBox cboWewenang 
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
         Left            =   1575
         TabIndex        =   2
         Top             =   855
         Width           =   4440
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
         TabIndex        =   0
         Top             =   195
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
         TabIndex        =   1
         Top             =   525
         Width           =   4440
      End
      Begin VB.Label Label11 
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
         Left            =   210
         TabIndex        =   25
         Top             =   570
         Width           =   285
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pegawai"
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
         Left            =   210
         TabIndex        =   24
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wewenang"
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
         Left            =   180
         TabIndex        =   23
         Top             =   900
         Width           =   945
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   -15
      TabIndex        =   13
      Top             =   1125
      Width           =   6165
      Begin VB.CheckBox chTampil 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tampil Karakter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1590
         TabIndex        =   4
         Top             =   615
         Width           =   4410
      End
      Begin VB.TextBox trPas 
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
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1215
         Width           =   2940
      End
      Begin VB.TextBox tPas 
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
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   885
         Width           =   2940
      End
      Begin VB.TextBox tUser 
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
         Top             =   210
         Width           =   4410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Password"
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
         Left            =   210
         TabIndex        =   18
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   210
         TabIndex        =   15
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   210
         TabIndex        =   14
         Top             =   900
         Width           =   765
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
      Left            =   0
      TabIndex        =   19
      Top             =   2640
      Width           =   6165
      Begin MSComctlLib.ListView vAnak 
         Height          =   3720
         Left            =   90
         TabIndex        =   12
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "a"
            Text            =   "No"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "b"
            Text            =   "USER NAME"
            Object.Width           =   6272
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "PASSWORD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "NAMA PEGAWAI"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "NIP"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "WEWENANG"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   900
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin VB.Label LPas 
      Caption         =   "Label1"
      Height          =   300
      Left            =   240
      TabIndex        =   17
      Top             =   3075
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LLevel 
      Caption         =   "Label3"
      Height          =   285
      Left            =   270
      TabIndex        =   16
      Top             =   3510
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cEdit

Private Sub cboWewenang_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub cboWewenang_LostFocus()
On Error Resume Next

cboWewenang.Text = Rep(cboWewenang.Text)
For i = 0 To cboWewenang.ListCount - 1
        If (UCase(cboWewenang.List(i)) Like "*" + UCase(cboWewenang.Text) + "*" = True) Then
            cboWewenang.Text = cboWewenang.List(i)
            Exit Sub
        End If
          If i = cboWewenang.ListCount - 1 Then
            If UCase(cboWewenang.List(i)) Like "*" + UCase(cboWewenang.Text) + "*" = False Then
                cboWewenang.Text = cboWewenang.List(0)
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub chTampil_Click()
On Error Resume Next
If chTampil.Value = 0 Then
    tPas.PasswordChar = "*"
    trPas.PasswordChar = "*"
    chTampil.Caption = "Sembunyikan Karakter"
Else
    tPas.PasswordChar = ""
    trPas.PasswordChar = ""
    chTampil.Caption = "Tampilkan Karakter"
End If
End Sub

Private Sub chTampil_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

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
If UCase(tPas.Text) <> UCase(trPas.Text) Then
    MsgBox "Konfirmasi Pasword tidak sama", vbCritical, "Error"
    trPas.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
End If
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
'MsgBox cEdit
'C_STR = "Select * from [USERs] where ([UserName]+'-'+[NIP]+'-'+WEWENANG)='" & txID.Text & "' order by [USERNAME] asc"
'openDB (C_STR)
'If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'If Not rPajak.EOF Then
    If cEdit = 1 Then
        E_STR = "UPDATE [USERS] SET [USERNAME]='" & tUser.Text & "',[PASSWORD]='" & tPas.Text & "',[NAMA]='" & tNama.Text & "',[NIP]='" & tNIP.Text & "',[WEWENANG]='" & Left(Trim(cboWewenang.Text), 2) & "' WHERE  ([USERNAME]+'-'+[NIP]+'-'+WEWENANG)='" & txID.Text & "' "
        'E_STR = "UPDATE [USERS] SET ([USERNAME]='" & tUser.Text & "',[PASSWORD]='" & tPas.Text & "',[NAMA]='" & tNama.Text & "',[NIP]='" & tNIP.Text & "',[WEWENANG]='" & Left(Trim(cboWewenang.Text), 2) & "',[ID]='" & txID & "') WHERE  [ID]='" & txID.Text & "' "
        
        openDB (E_STR)
         TampilUser
        Lama
        Screen.MousePointer = vbDefault
'        Exit Sub
'    Else
'        MsgBox "User sudah terdaftar, silahkan diganti", vbCritical, "Error"
'        tUser.SetFocus
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'End If
'If cEdit = 0 Then
    Else
        'MsgBox "TES"
        I_STR = "INSERT INTO [USERS] ([USERNAME],[PASSWORD],NAMA,NIP,WEWENANG) VALUES ('" & tUser.Text & "','" & tPas.Text & "','" & tNama.Text & "','" & tNIP.Text & "','" & Left(Trim(cboWewenang.Text), 2) & "')"
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
On Error Resume Next
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
c_Wewenang
chTampil.Value = 0
chTampil.Caption = "Tampilkan Karakter"
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
chTampil.Enabled = True
cmdSimpan.Caption = "&Simpan"
cmdBaru.Caption = "&Baru"
tNama.SetFocus
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
chTampil.Enabled = False
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
If InStr("0123456789. ", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub tPas_GotFocus()
On Error Resume Next
tPas.SelStart = 0
tPas.SelLength = Len(tPas.Text)
End Sub

Private Sub tPas_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
     KeyAscii = 0
End If
End Sub

Private Sub tPas_LostFocus()
tPas.Text = Rep(tPas.Text)
End Sub

Private Sub trPas_GotFocus()
On Error Resume Next
trPas.SelStart = 0
trPas.SelLength = Len(trPas.Text)
End Sub

Private Sub trPas_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
     KeyAscii = 0
End If

End Sub

Private Sub trPas_LostFocus()
trPas.Text = Rep(trPas.Text)
End Sub

Private Sub tUser_GotFocus()
On Error Resume Next
tUser.SelStart = 0
tUser.SelLength = Len(tUser.Text)
End Sub

Private Sub tUser_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
     KeyAscii = 0
End If
End Sub

Private Sub tUser_LostFocus()
tUser.Text = Rep(tUser.Text)
End Sub

Private Sub vAnak_Click()
On Error Resume Next
  tNama.Text = vAnak.SelectedItem.ListSubItems(3).Text
  tNIP = vAnak.SelectedItem.ListSubItems(4).Text
  cboWewenang.Text = vAnak.SelectedItem.ListSubItems(5).Text
  tUser.Text = vAnak.SelectedItem.ListSubItems(1).Text
  tPas.Text = vAnak.SelectedItem.ListSubItems(2).Text
  trPas.Text = vAnak.SelectedItem.ListSubItems(2).Text
  txID.Text = vAnak.SelectedItem.ListSubItems(6).Text
  For i = 0 To cboWewenang.ListCount - 1
        If (UCase(cboWewenang.List(i)) Like "*" + UCase(cboWewenang.Text) + "*" = True) Then
            cboWewenang.Text = cboWewenang.List(i)
            Exit Sub
        End If
          If i = cboWewenang.ListCount - 1 Then
            If UCase(cboWewenang.List(i)) Like "*" + UCase(cboWewenang.Text) + "*" = False Then
                cboWewenang.Text = cboWewenang.List(0)
                Exit Sub
            End If
        End If
    Next
End Sub
Sub TampilUser()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
vAnak.ListItems.Clear
C_STR = "select * from [USERS]"
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'rsAdmin.Find "NIP= '" & tNIP.Text & "'"
i = 0
Do While Not rPajak.EOF
    i = i + 1
    vAnak.ListItems.Add i, "", Format(i, "000")
    vAnak.ListItems.Item(i).ListSubItems.Add 1, "", rPajak![UserName]
    vAnak.ListItems.Item(i).ListSubItems.Add 2, "", rPajak![Password]
    vAnak.ListItems.Item(i).ListSubItems.Add 3, "", rPajak![NAMA]
    vAnak.ListItems.Item(i).ListSubItems.Add 4, "", rPajak![NIP]
    vAnak.ListItems.Item(i).ListSubItems.Add 5, "", rPajak![WEWENANG]
    vAnak.ListItems.Item(i).ListSubItems.Add 6, "", rPajak![UserName] & "-" & rPajak![NIP] & "-" & rPajak!WEWENANG
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
C_SSTR = "select * from USERS"
openDB (C_SSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If rPajak.RecordCount = 1 Then MsgBox "User tidak dapat dihapus...!", vbCritical, "Error": Screen.MousePointer = vbDefault: Exit Sub
C_STR = "DELETE From [USERS] where [USERNAME]+'-'+[NIP]+'-'+WEWENANG ='" & txID.Text & "'"
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
Sub c_Wewenang()
On Error GoTo Salah
cboWewenang.Clear
C_STR = "select * from WEWENANG"
openDB (C_STR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    cboWewenang.AddItem rPajak!KD_WEWENANG & "-" & rPajak!NM_WEWENANG
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub vAnak_KeyDown(KeyCode As Integer, Shift As Integer)
vAnak_Click
End Sub

Private Sub vAnak_KeyUp(KeyCode As Integer, Shift As Integer)
vAnak_Click
End Sub
