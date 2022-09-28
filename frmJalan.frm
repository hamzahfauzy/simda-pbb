VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJalan 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Nama Jalan/Dusun"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   12000
   Begin VB.Frame Frame4 
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
      Height          =   525
      Left            =   8715
      TabIndex        =   20
      Top             =   -60
      Visible         =   0   'False
      Width           =   3210
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
         Left            =   1185
         TabIndex        =   3
         Top             =   150
         Width           =   1935
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
         Height          =   225
         Left            =   60
         TabIndex        =   21
         Top             =   225
         Width           =   1215
      End
   End
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
      Height          =   525
      Left            =   45
      TabIndex        =   19
      Top             =   -60
      Width           =   11895
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
         Left            =   8865
         TabIndex        =   2
         Top             =   210
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
         Left            =   330
         TabIndex        =   0
         Top             =   210
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
         Left            =   4755
         TabIndex        =   1
         Top             =   210
         Width           =   1695
      End
   End
   Begin VB.ComboBox ccT 
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
      Left            =   645
      TabIndex        =   18
      Top             =   7200
      Visible         =   0   'False
      Width           =   1680
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
      Left            =   6345
      TabIndex        =   11
      Top             =   7170
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
      Left            =   5445
      TabIndex        =   10
      Top             =   7170
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
      Left            =   4545
      TabIndex        =   9
      Top             =   7170
      Width           =   915
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
      Height          =   915
      Left            =   30
      TabIndex        =   12
      Top             =   360
      Width           =   11910
      Begin VB.TextBox tJalan 
         Alignment       =   2  'Center
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
         Left            =   6645
         TabIndex        =   6
         Top             =   180
         Width           =   5160
      End
      Begin VB.TextBox tSem 
         Alignment       =   2  'Center
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
         Left            =   1485
         TabIndex        =   22
         Top             =   750
         Visible         =   0   'False
         Width           =   4380
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
         ItemData        =   "frmJalan.frx":0000
         Left            =   1230
         List            =   "frmJalan.frx":0002
         TabIndex        =   5
         Top             =   510
         Width           =   4380
      End
      Begin VB.ComboBox ccKec 
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
         Left            =   1230
         TabIndex        =   4
         Text            =   "ccKec"
         Top             =   165
         Width           =   4380
      End
      Begin VB.TextBox Text9 
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
         Left            =   12570
         TabIndex        =   13
         Top             =   315
         Width           =   4440
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6645
         TabIndex        =   23
         Top             =   390
         Width           =   5160
         Begin VB.TextBox tKelas 
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
            Left            =   10020
            TabIndex        =   25
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox tNIR 
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
            Left            =   6645
            TabIndex        =   24
            Top             =   255
            Width           =   2280
         End
         Begin VB.ComboBox ccZNT 
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
            Left            =   3120
            TabIndex        =   8
            Top             =   120
            Width           =   1995
         End
         Begin VB.ComboBox ccBlok 
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
            Left            =   540
            TabIndex        =   7
            Top             =   120
            Width           =   1605
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Kelas Bumi"
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
            Left            =   9135
            TabIndex        =   29
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode ZNT"
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
            Left            =   2340
            TabIndex        =   28
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label5 
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
            Height          =   165
            Left            =   75
            TabIndex        =   27
            Top             =   210
            Width           =   1215
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "NIR"
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
            Left            =   6150
            TabIndex        =   26
            Top             =   300
            Width           =   540
         End
      End
      Begin VB.Label Label4 
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
         TabIndex        =   17
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         TabIndex        =   15
         Top             =   555
         Width           =   1665
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Jalan"
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
         Left            =   5760
         TabIndex        =   14
         Top             =   210
         Width           =   1320
      End
   End
   Begin MSComctlLib.ListView vBangunan 
      Height          =   5805
      Left            =   30
      TabIndex        =   16
      Top             =   1260
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   10239
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
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
         Text            =   "BLOK"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NAMA JALAN"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "ZNT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "NIR"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "KELAS"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "NJOP Bumi/M2"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "KEL_MIN"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "KEL_MAX"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "KEC"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "KEL"
         Object.Width           =   1411
      EndProperty
   End
End
Attribute VB_Name = "frmJalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ccProses

Private Sub ccBlok_Change()
'CALL_ZNT
'CALL_JALAN2
End Sub

Private Sub ccBlok_Click()
On Error Resume Next
CALL_ZNT
If zJalan = 1 Then
    ZJALAN_1
    CALL_NIR1
    CALL_KELAS1
Else
    CALL_JALAN2
End If
End Sub

Private Sub ccBlok_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub ccBlok_LostFocus()
On Error Resume Next
ccBlok.Text = Rep(ccBlok.Text)
For i = 0 To ccBlok.ListCount - 1
        If (UCase(ccBlok.List(i)) Like "*" + UCase(ccBlok.Text) + "*" = True) Then
            ccBlok.Text = ccBlok.List(i)
            ccBlok_Click
            Exit Sub
        End If
          If i = ccBlok.ListCount - 1 Then
            If UCase(ccBlok.List(i)) Like "*" + UCase(ccBlok.Text) + "*" = False Then
                ccBlok.Text = ccBlok.List(0)
                ccBlok_Click
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub ccKec_Click()
On Error Resume Next
CALL_KEL
CALL_Jalan1
CALL_BLOK
If zJalan = 1 Then CALL_NIR1: CALL_KELAS1

End Sub

Private Sub ccKec_GotFocus()
On Error Resume Next
ccProses = 1
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
'CALL_JALAN
'CALL_BUMI
On Error Resume Next
CALL_Jalan1
CALL_BLOK
If zJalan = 1 Then CALL_NIR1: CALL_KELAS1
End Sub

Private Sub ccKel_GotFocus()
On Error Resume Next
ccProses = 2
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

Private Sub ccZNT_Click()
'callNIR
'CALL_KELAS
On Error Resume Next
If zJalan = 1 Then
    ZJALAN_1
    CALL_NIR1
    CALL_KELAS1
Else
    ZJALAN_1
'CALL_JALAN2
End If
End Sub

Private Sub ccZNT_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Private Sub ccZNT_LostFocus()
On Error Resume Next
ccZNT.Text = Rep(ccZNT.Text)
For i = 0 To ccZNT.ListCount - 1
        If (UCase(ccZNT.List(i)) Like "*" + UCase(ccZNT.Text) + "*" = True) Then
            ccZNT.Text = ccZNT.List(i)
            ccZNT_Click
            Exit Sub
        End If
          If i = ccZNT.ListCount - 1 Then
            If UCase(ccZNT.List(i)) Like "*" + UCase(ccZNT.Text) + "*" = False Then
                ccZNT.Text = ccZNT.List(0)
                ccZNT_Click
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
On Error Resume Next
ccBlok.Text = ""
ccZNT.Text = ""
tNIR.Text = 0
TKELAS.Text = 0
ccKec.Text = ""
ccKel.Text = ""
ccBlok.Text = ""
ccZNT.Text = ""
tJalan.Text = ""
tSem.Text = ""
vBangunan.ListItems.Clear
CALL_KEC
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
Unload Me
zJalan = ""
End Sub

Private Sub cmdSave_Click()
On Error GoTo Salah
If zJalan = 1 Then
    frmObjek_Pajak_Bm.cboNOP(1).Clear
    frmObjek_Pajak_Bm.cboNOP(2).Clear
  frmObjek_Pajak_Bm.cboJalan.Text = vBangunan.SelectedItem.ListSubItems(4).Text & "-" & vBangunan.SelectedItem.ListSubItems(3).Text
    frmObjek_Pajak_Bm.cboNOP(0).Text = Trim(ccKec.Text) 'Left(Trim(ccKec.Text), 3) & "-" & Mid(Trim(ccKec.Text), 5, Len(ccKec.Text))
    frmObjek_Pajak_Bm.cboNOP(1).Text = Trim(ccKel.Text) 'Left(Trim(ccKel.Text), 3) & "-" & Mid(Trim(ccKel.Text), 5, Len(ccKel.Text))
    frmObjek_Pajak_Bm.cboNOP(2).Text = Trim(vBangunan.SelectedItem.ListSubItems(2).Text)
    frmObjek_Pajak_Bm.tBumi(6).Text = vBangunan.SelectedItem.ListSubItems(4).Text
    frmObjek_Pajak_Bm.cboNOP(1).SetFocus
    frmObjek_Pajak_Bm.Show
    Unload Me
    Exit Sub
    
End If
Dim CD_KEC, CD_KEL
CD_KEC = Left(Trim(ccKec.Text), 3)
CD_KEL = Left(Trim(ccKel.Text), 3)
B_SQL = "SELECT KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,JALAN_OP FROM DAT_OBJEK_PAJAK WHERE KD_KECAMATAN='" & CD_KEC & "'  AND KD_KELURAHAN='" & CD_KEL & "'  AND KD_BLOK='" & ccBlok.Text & "' AND JALAN_OP='" & tJalan.Text & "'"
openDB (B_SQL)
If rPajak.RecordCount > 0 Then rPajak.MoveNext
If Not rPajak.EOF And (chPajak(2).Value = 1 Or chPajak(3).Value = 1) Then
    MsgBox "ANDA TIDAK DAPAT MENGEDIT/HAPUS NAMA JALAN INI", vbCritical, "TETNONG"
    Exit Sub
End If
If ccKec.Text = "" Or ccKel.Text = "" Or ccBlok.Text = "" Or ccZNT.Text = "" Or tJalan.Text = "" Then
     MsgBox "MASIH ADA DATA YANG KOSONG...", vbCritical, "TETNONG"
       ' ccKec.SetFocus
            Exit Sub
End If

If chPajak(1).Value = 1 Then
xTanya = MsgBox("APA ANDA YAKIN MENYIMPAN DATA INI?", vbInformation + vbYesNo, "Info")
If xTanya = vbNo Then Exit Sub
    If Not rPajak.EOF Then
        MsgBox "Data Sudah Ada, Silahkan Diganti...", vbCritical, "Tetnong"
        Exit Sub
    End If
    iSQL = "Insert Into JALAN_STANDARD VALUES ('12','12','" & CD_KEC & "','" & CD_KEL & "','" & tJalan.Text & "','" & tJalan.Text & "')"
    openDB (iSQL)
    iSQL2 = "Insert Into JALAN VALUES ('12','12','" & CD_KEC & "','" & CD_KEL & "','" & ccBlok.Text & "','" & ccZNT.Text & "','" & tJalan.Text & "')"
    openDB (iSQL2)
ElseIf chPajak(2).Value = 1 Then
xTanya = MsgBox("APA ANDA YAKIN MENGHAPUS DATA INI?", vbInformation + vbYesNo, "Info")
If xTanya = vbNo Then Exit Sub

'    If rPajak.EOF Then
'        MsgBox "Data Tidak Ada...", vbCritical, "Tetnong"
'        Exit Sub
'    End If
    iSQL1 = "Delete From JALAN_STANDARD WHERE KD_KECAMATAN='" & CD_KEC & "'  AND KD_KELURAHAN='" & CD_KEL & "'  AND NM_JLN_SEMENTARA='" & tSem.Text & "'"
    openDB (iSQL1)
    iSQL2 = "Delete From JALAN WHERE KD_KECAMATAN='" & CD_KEC & "'  AND KD_KELURAHAN='" & CD_KEL & "'  AND KD_BLOK='" & ccBlok.Text & "' AND KD_ZNT='" & ccZNT.Text & "' AND NM_JLN='" & Trim(tSem.Text) & "'"
    openDB (iSQL2)
ElseIf chPajak(3).Value = 1 Then
xTanya = MsgBox("APA ANDA YAKIN MENGEDIT DATA INI?", vbInformation + vbYesNo, "Info")
If xTanya = vbNo Then Exit Sub

'If rPajak.RecordCount > 0 Then rPajak.MoveFirst
'
'    If rPajak.EOF Then
'        MsgBox "Data Tidak Ada...", vbCritical, "Tetnong"
'        Exit Sub
'    End If
    'rPajak!KD_PROPINSI = "12"
    'rPajak!KD_DATI2 = "12"
    'rPajak!KD_KECAMATAN = CD_KEC
    'rPajak!KD_KELURAHAN = CD_KEL
    'rPajak!KD_BLOK = ccBlok.Text
    'rPajak!KD_BLOK = ccBlok.Text
    'rPajak!NM_JLN = tJalan.Text
    'rPajak.Update
    'iSQL = "Select * From JALAN WHERE KD_KECAMATAN='" & CD_KEC & "'  AND KD_KELURAHAN='" & CD_KEL & "'  AND KD_BLOK='" & ccBlok.Text & "' AND KD_ZNT='" & ccZNT.Text & "' AND NM_JLN='" & tSem.Text & "'   ORDER BY NM_JLN_SEMENTARA,KD_KELURAHAN, KD_KECAMATAN ASC"
    iSQL = "update JALAN SET KD_PROPINSI='12',KD_DATI2='12',KD_KECAMATAN='" & CD_KEC & "',KD_KELURAHAN='" & CD_KEL & "' , KD_BLOK='" & ccBlok.Text & "', KD_ZNT='" & ccZNT.Text & "',NM_JLN='" & tJalan.Text & "' where (KD_KECAMATAN='" & CD_KEC & "' AND KD_KELURAHAN='" & CD_KEL & "' AND KD_BLOK='" & ccBlok.Text & "' AND NM_JLN='" & tSem.Text & "')"
    openDB (iSQL)
    iSQL1 = "UPDATE JALAN_STANDARD SET KD_PROPINSI='12',KD_DATI2='12',KD_KECAMATAN='" & CD_KEC & "',KD_KELURAHAN='" & CD_KEL & "'  where KD_KECAMATAN='" & CD_KEC & "' AND KD_KELURAHAN='" & CD_KEL & "' AND NM_JLN_SEMENTARA='" & tSem.Text & "'"
    'iSQL = "UPDATE REF_KECAMATAN SET KD_PROPINSI='12',KD_DATI2='12',KD_KECAMATAN='" & tKode.Text & "',NM_KECAMATAN='" & tKec.Text & "' where KD_KECAMATAN='" & tKode.Text & "'"
    openDB (iSQL1)
End If
'CALL_BLOK
CALL_Jalan1
CALL_BLOK

ccKec.Text = ""
ccKel.Text = ""
ccBlok.Text = ""
ccZNT.Text = ""
tJalan.Text = ""
tSem.Text = ""

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub Form_Activate()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
ccTahun.Clear
QSTR = "SELECT * FROM KELAS_TANAH ORDER BY THN_AWAL_KLS_TANAH ASC"
openDB (QSTR)
 ccTahun.Clear
ccTahun.Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    ccTahun.AddItem i
Next
ccT.Clear
QSTR = "SELECT THN_AWAL_KLS_TANAH FROM KELAS_TANAH GROUP BY THN_AWAL_KLS_TANAH ORDER BY THN_AWAL_KLS_TANAH DESC"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
   ccT.AddItem rPajak!THN_AWAL_KLS_TANAH
    rPajak.MoveNext
Loop
ccT.Text = ccT.List(0)
    cmdSave.Caption = "&Save"
    Label3.Visible = True
    tJalan.Visible = True
    Frame1.Visible = True
    Frame3.Visible = True
    Frame1.Top = 390
    Label4.Top = 195
    Label2.Top = 555
    ccKec.Top = 165
    ccKel.Top = 510
    Label2.Left = 150
    ccKel.Left = 1230
    vBangunan.ColumnHeaders(6).Width = 0
    vBangunan.ColumnHeaders(7).Width = 0
    vBangunan.ColumnHeaders(8).Width = 0
    vBangunan.ColumnHeaders(9).Width = 0
    vBangunan.ColumnHeaders(10).Width = 0
If zJalan = 1 Then
    ccKec.Text = ""
    ccKel.Text = ""
    ccBlok.Text = ""
    vBangunan.ColumnHeaders(6).Width = 1100
    vBangunan.ColumnHeaders(7).Width = 700
    vBangunan.ColumnHeaders(8).Width = 1100
    vBangunan.ColumnHeaders(9).Width = 1100
    vBangunan.ColumnHeaders(10).Width = 1100
    ccKec.Text = Trim(frmObjek_Pajak_Bm.cboNOP(0).Text)
    ccKel.Text = Trim(frmObjek_Pajak_Bm.cboNOP(1).Text)
    ccBlok.Text = frmObjek_Pajak_Bm.cboNOP(2).Text
    cmdSave.Caption = "&OK"
    Label3.Visible = False
    tJalan.Visible = False
    'Frame1.Visible = False
    Frame1.Top = 220
    Frame3.Visible = False
    'Label4.Top = 420
    'Label2.Top = 420
    'ccKec.Top = 390
    'ccKel.Top = 390
    'Label2.Left = 5760
    'ccKel.Left = 6645
    'ccZNT.Visible = False
    'Label6.Visible = False
    If Trim(ccKec.Text) <> "" And Trim(ccKel.Text) = "" And Trim(ccBlok.Text) = "" Then
        CALL_KEL
        CALL_BLOK
        ZJALAN_1
        CALL_NIR1
        CALL_KELAS1
        GoTo Salah
    ElseIf Trim(ccKec.Text) <> "" And Trim(ccKel.Text) <> "" Then
        CALL_BLOK
        ZJALAN_1
        CALL_NIR1
        CALL_KELAS1
        GoTo Salah
        
    End If
End If

CALL_KEC

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub

Sub CALL_KEC()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
ccKec.Clear
QSTR = "SELECT * FROM REF_KECAMATAN ORDER BY KD_KECAMATAN ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            ccKec.AddItem rPajak!KD_KECAMATAN & "-" & rPajak!NM_KECAMATAN
        rPajak.MoveNext
        Loop

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
QSTR = "SELECT * FROM REF_KELURAHAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' ORDER BY KD_KELURAHAN ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            ccKel.AddItem rPajak!KD_KELURAHAN & "-" & rPajak!NM_KELURAHAN
        rPajak.MoveNext
        Loop
    
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub
Sub CALL_Jalan()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
Dim i As Integer
vBangunan.ListItems.Clear
If ccProses = 1 Then
    QSTR = "SELECT * FROM JALAN_STANDARD WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' ORDER BY KD_KECAMATAN ASC"
Else
    QSTR = "SELECT * FROM JALAN_STANDARD WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'  ORDER BY KD_KELURAHAN ASC"
End If
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "000")
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_KECAMATAN])
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", rPajak![KD_KELURAHAN]
        If rPajak![NM_JLN_SEMENTARA] = "" Or IsNull(rPajak![NM_JLN_SEMENTARA]) = True Then
            vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", "-"
        Else
            vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", rPajak![NM_JLN_SEMENTARA]
        End If
        If rPajak![NM_JLN_STANDARD] = "" Or IsNull(rPajak![NM_JLN_STANDARD]) = True Then
            vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", "-"
        Else
            vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", rPajak![NM_JLN_STANDARD]
        End If
                
        rPajak.MoveNext
        Loop

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub


Private Sub tJalan_GotFocus()
On Error Resume Next
tJalan.Alignment = 0
End Sub

Private Sub tJalan_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    KeyAscii = 0
End If
End Sub

Private Sub tJalan_LostFocus()
On Error Resume Next
tJalan.Alignment = 2
tJalan.Text = Rep(tJalan.Text)
End Sub

Private Sub tKelas_GotFocus()
On Error Resume Next
TKELAS.Alignment = 0
End Sub

Private Sub tKelas_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub tKelas_LostFocus()
On Error Resume Next
TKELAS.Alignment = 1
End Sub

Private Sub tNIR_GotFocus()
On Error Resume Next
tNIR.Alignment = 0
End Sub

Private Sub tNIR_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub tNIR_LostFocus()
On Error Resume Next
tNIR.Alignment = 1
CALL_KELAS

End Sub

Private Sub vBangunan_Click()
On Error GoTo Salah
If chPajak(2).Value = 1 Or chPajak(3).Value = 1 Or zJalan = 1 Then
'    tBlok.Text = vBangunan.SelectedItem.ListSubItems(4).Text
    For i = 0 To ccKec.ListCount - 1
        If vBangunan.SelectedItem.ListSubItems(10).Text = Left(Trim(ccKec.List(i)), 3) Then
           ccKec.Text = ccKec.List(i)
        End If
    Next
    For i = 0 To ccKel.ListCount - 1
        If vBangunan.SelectedItem.ListSubItems(11).Text = Left(Trim(ccKel.List(i)), 3) Then
           ccKel.Text = ccKel.List(i)
        End If
    Next
    ccBlok.Text = vBangunan.SelectedItem.ListSubItems(2).Text
    tJalan.Text = vBangunan.SelectedItem.ListSubItems(3).Text
    tSem.Text = vBangunan.SelectedItem.ListSubItems(3).Text
    ccZNT.Text = vBangunan.SelectedItem.ListSubItems(4).Text
End If

If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub vBangunan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
vBangunan.SortKey = ColumnHeader.Index - 1
vBangunan.Sorted = True
vBangunan.Sorted = False
vBangunan.SortOrder = lvwAscending

End Sub

Sub CALL_BLOK()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
ccBlok.Clear
If ccProses = 1 Then
    QSTR = "SELECT KD_BLOK FROM DAT_PETA_BLOK WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' GROUP BY KD_BLOK ORDER BY KD_BLOK ASC"
Else
    QSTR = "SELECT * FROM DAT_PETA_BLOK WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' ORDER BY KD_BLOK,KD_KELURAHAN ASC"
End If

openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            ccBlok.AddItem rPajak!KD_BLOK
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
Screen.MousePointer = vbDefault
End Sub

Sub CALL_ZNT()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
ccZNT.Clear
If ccProses = 1 Then
    'QSTR = "SELECT * FROM DAT_PETA_ZNT WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_BLOK='" & ccBlok.Text & "' ORDER BY KD_ZNT,KD_KECAMATAN ASC"
    
    QSTR = "SELECT KD_ZNT,KD_KECAMATAN FROM DAT_PETA_ZNT WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_BLOK='" & ccBlok.Text & "' GROUP BY KD_ZNT,KD_KECAMATAN ORDER BY KD_ZNT,KD_KECAMATAN ASC"
Else
    QSTR = "SELECT * FROM DAT_PETA_ZNT WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' AND KD_BLOK='" & ccBlok.Text & "' ORDER BY KD_ZNT,KD_KELURAHAN ASC"
End If
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            ccZNT.AddItem rPajak!KD_ZNT
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
Screen.MousePointer = vbDefault
End Sub
Sub CALL_KELAS()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
TKELAS.Text = 0
QSTR = "SELECT * FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH ='" & Trim(ccT.Text) & "' ORDER BY THN_AWAL_KLS_TANAH ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            If tNIR.Text = "" Then tNIR.Text = 0
            If tNIR.Text * 1 >= rPajak!NILAI_MIN_TANAH * 1000 And tNIR.Text * 1 <= rPajak!NILAI_MAX_TANAH * 1000 Then
                TKELAS.Text = rPajak!KD_KLS_TANAH
            End If
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
Screen.MousePointer = vbDefault
End Sub

Sub callNIR()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
tNIR.Text = 0
If ccProses = 1 Then
    strKab = "Select * From DAT_NIR where KD_ZNT = '" & ccZNT.Text & "' and THN_NIR_ZNT='" & Trim(ccTahun.Text) - 1 & "' and KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' order by THN_NIR_ZNT,KD_ZNT asc"
Else
    strKab = "Select * From DAT_NIR where KD_ZNT = '" & ccZNT.Text & "' and THN_NIR_ZNT='" & Trim(ccTahun.Text) - 1 & "' and KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' and KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' order by THN_NIR_ZNT,KD_ZNT asc"
End If
openDB (strKab)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    'tBumi(2).Text = Format(Trim(rPajak!NIR) * 1000, "#,#0")
    tNIR.Text = Format(Trim(rPajak!NIR) * 1000, "#,#0.00")
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub
'Sub CALL_BUMI()
'Dim I As Integer
'vBangunan.ListItems.Clear
'QSTR = "SELECT * FROM KELAS_TANAH WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'  ORDER BY BLOK,KD_ZNT ASC"
'BUKA1 (QSTR)
' If rBUMI.RecordCount > 0 Then rBUMI.MoveFirst
'        I = 0
'        Do While Not rBUMI.EOF
'        I = I + 1
'        vBangunan.ListItems.Add I, "", Format(I, "#0")
'        vBangunan.ListItems.Item(I).ListSubItems.Add 1, "", Format(I, "000")
'        'vBangunan.ListItems.Item(I).ListSubItems.Add 2, "", Trim(rBUMI![KD_KECAMATAN])
'        vBangunan.ListItems.Item(I).ListSubItems.Add 2, "", Trim(rBUMI![BLOK])
'        vBangunan.ListItems.Item(I).ListSubItems.Add 3, "", rBUMI![JALAN_OP]
'        vBangunan.ListItems.Item(I).ListSubItems.Add 4, "", rBUMI![KD_ZNT]
'        vBangunan.ListItems.Item(I).ListSubItems.Add 5, "", Format(rBUMI![NIR], "#,#0.00")
'        vBangunan.ListItems.Item(I).ListSubItems.Add 6, "", rBUMI![KELAS]
'        vBangunan.ListItems.Item(I).ListSubItems.Add 7, "", Format(rBUMI![NJOP], "#,#0.00")
'        vBangunan.ListItems.Item(I).ListSubItems.Add 8, "", Format(rBUMI![NILAI_MIN_TANAH], "#,#0.00")
'        vBangunan.ListItems.Item(I).ListSubItems.Add 9, "", Format(rBUMI![NILAI_MAX_TANAH], "#,#0.00")
'
'        rBUMI.MoveNext
'        Loop
'End Sub
Sub CALL_Jalan1()
On Error GoTo Salah
Dim i As Integer
vBangunan.ListItems.Clear
If ccProses = 1 Then
    QSTR = "SELECT * FROM JALAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' ORDER BY KD_BLOK,KD_ZNT ASC"
Else
    QSTR = "SELECT * FROM JALAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'  ORDER BY KD_BLOK,KD_ZNT ASC"
End If
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "000")
        'vBangunan.ListItems.Item(I).ListSubItems.Add 2, "", Trim(rBUMI![KD_KECAMATAN])
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_BLOK])
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", rPajak![NM_JLN]
        vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", rPajak![KD_ZNT]
        vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", Trim(rPajak![KD_KECAMATAN])
        vBangunan.ListItems.Item(i).ListSubItems.Add 11, "", Trim(rPajak![KD_KELURAHAN])
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Sub CALL_KELAS1()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH='" & ccT.Text & "'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    For J = 1 To vBangunan.ListItems.Count
    If vBangunan.ListItems.Item(J).ListSubItems(5).Text * 1 >= rPajak!NILAI_MIN_TANAH * 1000 And vBangunan.ListItems.Item(J).ListSubItems(5).Text * 1 <= rPajak!NILAI_MAX_TANAH * 1000 Then
        vBangunan.ListItems.Item(J).ListSubItems(6).Text = Format(rPajak!KD_KLS_TANAH)
        vBangunan.ListItems.Item(J).ListSubItems(7).Text = Format(rPajak!NILAI_PER_M2_TANAH * 1000, "#,#0.00")
        vBangunan.ListItems.Item(J).ListSubItems(8).Text = Format(rPajak!NILAI_MIN_TANAH * 1000, "#,#0.00")
        vBangunan.ListItems.Item(J).ListSubItems(9).Text = Format(rPajak!NILAI_MAX_TANAH * 1000, "#,#0.00")
        vBangunan.ListItems.Item(J).ListSubItems(7).ForeColor = vbBlue
    End If
    Next
rPajak.MoveNext

Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Sub CALL_NIR1()
'tNIR.Text = 0
On Error GoTo Salah
Screen.MousePointer = vbHourglass

If ccProses = 1 Then
    strKab = "Select * From DAT_NIR where THN_NIR_ZNT='" & Trim(ccTahun.Text * 1) - 1 & "' and KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' order by KD_ZNT asc"
Else
    strKab = "Select * From DAT_NIR where THN_NIR_ZNT='" & Trim(ccTahun.Text * 1) - 1 & "' and KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' and KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' order by KD_ZNT asc"
End If
openDB (strKab)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
Do While Not rPajak.EOF
    For i = 1 To vBangunan.ListItems.Count
        If UCase(Trim(rPajak!KD_ZNT)) = UCase(Trim(vBangunan.ListItems.Item(i).ListSubItems(4).Text)) Then
            vBangunan.ListItems.Item(i).ListSubItems(5).Text = Format(rPajak!NIR * 1000, "#,#0.00")
            
        End If
    Next
rPajak.MoveNext
Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub
Sub CALL_JALAN2()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
Dim i As Integer
vBangunan.ListItems.Clear
If ccProses = 1 Then
    QSTR = "SELECT * FROM JALAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_BLOK='" & Trim(ccBlok.Text) & "'  ORDER BY KD_BLOK,KD_ZNT ASC"
Else
    QSTR = "SELECT * FROM JALAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "'  AND KD_BLOK='" & Trim(ccBlok.Text) & "'  ORDER BY KD_BLOK,KD_ZNT ASC"
End If
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "000")
        'vBangunan.ListItems.Item(I).ListSubItems.Add 2, "", Trim(rBUMI![KD_KECAMATAN])
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_BLOK])
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", rPajak![NM_JLN]
        vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", rPajak![KD_ZNT]
        vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", Trim(rPajak![KD_KECAMATAN])
        vBangunan.ListItems.Item(i).ListSubItems.Add 11, "", Trim(rPajak![KD_KELURAHAN])
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description

Screen.MousePointer = vbDefault
End Sub

Private Sub vBangunan_DblClick()
On Error Resume Next
If zJalan = 1 Then
    frmObjek_Pajak_Bm.cboNOP(1).Clear
    frmObjek_Pajak_Bm.cboNOP(2).Clear
    frmObjek_Pajak_Bm.cboJalan.Text = vBangunan.SelectedItem.ListSubItems(4).Text & "-" & vBangunan.SelectedItem.ListSubItems(3).Text
    frmObjek_Pajak_Bm.cboNOP(0).Text = Trim(ccKec.Text) ' Left(Trim(ccKec.Text), 3) & "-" & Mid(Trim(ccKec.Text), 5, Len(ccKec.Text))
    frmObjek_Pajak_Bm.cboNOP(1).Text = Trim(ccKel.Text) 'Left(Trim(ccKel.Text), 3) & "-" & Mid(Trim(ccKel.Text), 5, Len(ccKel.Text))
    frmObjek_Pajak_Bm.cboNOP(2).Text = Trim(vBangunan.SelectedItem.ListSubItems(2).Text)
    frmObjek_Pajak_Bm.tBumi(6).Text = vBangunan.SelectedItem.ListSubItems(4).Text
    frmObjek_Pajak_Bm.cboNOP(1).SetFocus
    
    Unload Me
   
    frmObjek_Pajak_Bm.Show
    'MsgBox frmObjek_Pajak_Bm.cboJalan.Text
    
End If
End Sub

Sub ZJALAN_1()
On Error GoTo Salah
Dim i As Integer
vBangunan.ListItems.Clear
If Trim(ccKec.Text) <> "" And Trim(ccKel.Text) <> "" And Trim(ccBlok.Text) <> "" And Trim(ccZNT.Text) <> "" Then
    QSTR = "SELECT * FROM JALAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' AND KD_BLOK='" & Trim(ccBlok.Text) & "' AND KD_ZNT='" & Trim(ccZNT.Text) & "'  ORDER BY KD_BLOK,KD_ZNT ASC"
ElseIf Trim(ccKec.Text) <> "" And Trim(ccKel.Text) <> "" And Trim(ccBlok.Text) <> "" And Trim(ccZNT.Text) = "" Then
    QSTR = "SELECT * FROM JALAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' AND KD_BLOK='" & Trim(ccBlok.Text) & "' ORDER BY KD_BLOK,KD_ZNT ASC"
ElseIf Trim(ccKec.Text) <> "" And Trim(ccKel.Text) <> "" And Trim(ccBlok.Text) = "" And Trim(ccZNT.Text) = "" Then
    QSTR = "SELECT * FROM JALAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' ORDER BY KD_BLOK,KD_ZNT ASC"
ElseIf Trim(ccKec.Text) <> "" And Trim(ccKel.Text) = "" And Trim(ccBlok.Text) = "" And Trim(ccZNT.Text) = "" Then
    QSTR = "SELECT * FROM JALAN WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' ORDER BY KD_BLOK,KD_ZNT ASC"
End If
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#0")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "000")
        'vBangunan.ListItems.Item(I).ListSubItems.Add 2, "", Trim(rBUMI![KD_KECAMATAN])
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_BLOK])
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", rPajak![NM_JLN]
        vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", rPajak![KD_ZNT]
        vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", Trim(rPajak![KD_KECAMATAN])
        vBangunan.ListItems.Item(i).ListSubItems.Add 11, "", Trim(rPajak![KD_KELURAHAN])
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

