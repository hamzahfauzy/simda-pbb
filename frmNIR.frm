VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNIR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Nilai Indikasi Rata-Rata (NIR)"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12930
   ControlBox      =   0   'False
   Icon            =   "frmNIR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   12930
   Begin MSComctlLib.ListView vBangunan 
      Height          =   5760
      Left            =   60
      TabIndex        =   12
      Top             =   1620
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   10160
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
      NumItems        =   13
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
         Text            =   "ZNT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "NIR LAMA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "NIR BARU"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ket"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "KELAS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "NJOP"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "MIN"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "MAX"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "No. Dokumen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "KEL"
         Object.Width           =   2540
      EndProperty
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
      Height          =   615
      Left            =   8790
      TabIndex        =   24
      Top             =   -45
      Width           =   4080
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
         ItemData        =   "frmNIR.frx":0442
         Left            =   1695
         List            =   "frmNIR.frx":0444
         TabIndex        =   3
         Top             =   180
         Width           =   2250
      End
      Begin VB.Label Label5 
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
         Height          =   210
         Left            =   315
         TabIndex        =   25
         Top             =   255
         Width           =   1215
      End
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
      Left            =   60
      TabIndex        =   19
      Top             =   -45
      Width           =   8730
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
         Left            =   3240
         TabIndex        =   1
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
         Left            =   330
         TabIndex        =   0
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
         Left            =   6360
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
      Left            =   6570
      TabIndex        =   15
      Top             =   7545
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
      Left            =   5670
      TabIndex        =   14
      Top             =   7545
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
      Left            =   4770
      TabIndex        =   13
      Top             =   7545
      Width           =   915
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
      Height          =   615
      Left            =   60
      TabIndex        =   16
      Top             =   540
      Width           =   12810
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
         ItemData        =   "frmNIR.frx":0446
         Left            =   8100
         List            =   "frmNIR.frx":0448
         TabIndex        =   5
         Top             =   195
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
         ItemData        =   "frmNIR.frx":044A
         Left            =   990
         List            =   "frmNIR.frx":044C
         TabIndex        =   4
         Top             =   180
         Width           =   5205
      End
      Begin VB.Label Label3 
         Caption         =   "Kode Kelurahan"
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
         Left            =   6405
         TabIndex        =   18
         Top             =   210
         Width           =   1320
      End
      Begin VB.Label Label1 
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
         Left            =   90
         TabIndex        =   17
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   540
      Left            =   10305
      TabIndex        =   20
      Top             =   1065
      Width           =   2550
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
         TabIndex        =   10
         Text            =   "0"
         Top             =   165
         Width           =   960
      End
      Begin VB.CommandButton cmdPersen 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Picture         =   "frmNIR.frx":044E
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
         TabIndex        =   21
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Height          =   540
      Left            =   75
      TabIndex        =   22
      Top             =   1065
      Width           =   10230
      Begin VB.TextBox tDok 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   975
         TabIndex        =   6
         Top             =   120
         Width           =   2790
      End
      Begin VB.ComboBox ccZona 
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
         ItemData        =   "frmNIR.frx":1118
         Left            =   4380
         List            =   "frmNIR.frx":111A
         TabIndex        =   7
         Top             =   150
         Width           =   1800
      End
      Begin VB.TextBox tNIR 
         Alignment       =   1  'Right Justify
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
         Left            =   6750
         TabIndex        =   8
         Text            =   "0"
         Top             =   165
         Width           =   2970
      End
      Begin VB.CommandButton cmdCari 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9720
         Picture         =   "frmNIR.frx":111C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZNT"
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
         Left            =   3930
         TabIndex        =   27
         Top             =   195
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Dok"
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
         Left            =   105
         TabIndex        =   26
         Top             =   210
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Left            =   6300
         TabIndex        =   23
         Top             =   195
         Width           =   285
      End
   End
   Begin VB.Image Image1 
      Height          =   9180
      Left            =   0
      Picture         =   "frmNIR.frx":2D93
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "frmNIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xxTahun, ccProses
Dim ccNIR
Private Sub ccKec_Click()
On Error Resume Next
CALL_KEL
CALL_ZNT
CALL_NIR
tNIR.Text = 0
End Sub

Private Sub ccKec_GotFocus()
ccProses = 1
End Sub

Private Sub ccKec_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    ccKel.SetFocus
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
CALL_ZNT
CALL_NIR
tNIR.Text = 0
End Sub

Private Sub ccKel_GotFocus()
ccProses = 2
End Sub

Private Sub ccKel_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    tDok.SetFocus
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
                'CALL_ZNT
                'CALL_NIR
                ccKel_Click
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub ccTahun_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    ccKec.SetFocus
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

Private Sub ccZona_Change()
'CALL_NIR
End Sub

Private Sub ccZona_Click()
'CALL_NIR1
On Error GoTo Salah
If ccZona.Text <> "" And chPajak(1).Value = 1 Then
        C_STR = "SELECT * FROM DAT_NIR WHERE THN_NIR_ZNT='" & ccTahun.Text & "' AND KD_ZNT='" & ccZona.Text & "'"
        openDB (C_STR)
        If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        If Not rPajak.EOF Then
            ccNIR = rPajak!NIR
        Else
            ccNIR = 0
        End If

        i = vBangunan.ListItems.Count + 1
        For J = 1 To vBangunan.ListItems.Count
            If ccZona.Text = vBangunan.ListItems.Item(J).ListSubItems(2).Text Then Exit Sub
            
        Next
        
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", ccZona.Text
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Format(Trim(ccNIR), "#,#0.00")
        vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", 0
        vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", "-"
          vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 11, "", vBangunan.ListItems.Item(i - 1).ListSubItems(11).Text
        vBangunan.ListItems.Item(i).ListSubItems.Add 12, "", Left(Trim(ccKel.Text), 3)
End If
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub ccZona_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    tNIR.SetFocus
End If

End Sub

Private Sub ccZona_LostFocus()
On Error Resume Next
ccZona.Text = Rep(ccZona.Text)
For i = 0 To ccZona.ListCount - 1
        If (UCase(ccZona.List(i)) Like "*" + UCase(ccZona.Text) + "*" = True) Then
            ccZona.Text = ccZona.List(i)
           ccZona_Click ' CALL_NIR1
            Exit Sub
        End If
          If i = ccZona.ListCount - 1 Then
            If UCase(ccZona.List(i)) Like "*" + UCase(ccZona.Text) + "*" = False Then
                ccZona.Text = ccZona.List(0)
             ccZona_Click '   CALL_NIR1
                Exit Sub
            End If
        End If
    Next
    
End Sub

Private Sub chPajak_Click(Index As Integer)
On Error GoTo Salah
Select Case Index
Case 1
    If chPajak(1).Value = 1 Then
        chPajak(2).Value = 0
        chPajak(3).Value = 0
        cmdSave.Caption = "&Save"
        xxTahun = (ccTahun.Text * 1) - 1
        bersih
    End If
Case 2
    
   If chPajak(2).Value = 1 Then
    chPajak(1).Value = 0
    chPajak(3).Value = 0
    cmdSave.Caption = "&Delete"
    xTanya = MsgBox("Apa anda yakin menghapus NIR?", vbQuestion + vbYesNo, "Penghapusan")
    
    If xTanya = vbYes Then
        cmdSave.Caption = "&Delete"
        xxTahun = ccTahun.Text * 1
        bersih
    Else
        chPajak(1).Value = 1
        chPajak(2).Value = 0
        cmdSave.Caption = "&Save"
    End If
    
   End If
Case 3
    If chPajak(3).Value = 1 Then
        chPajak(1).Value = 0
        chPajak(2).Value = 0
        cmdSave.Caption = "&Update"
        xTanya = MsgBox("Apa anda yakin melakukan pemutakhiran NIR?", vbQuestion + vbYesNo, "Pemutakhiran")
        If xTanya = vbYes Then
            cmdSave.Caption = "&Update"
            xxTahun = ccTahun.Text * 1
            bersih
        Else
            chPajak(1).Value = 1
            chPajak(3).Value = 0
            cmdSave.Caption = "&Save"
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
    SendKeys "{tab}"
End If
End Sub

Private Sub cmdCari_Click()
On Error GoTo Salah
vBangunan.SelectedItem.ListSubItems(4).Text = tNIR.Text
vBangunan.SelectedItem.ListSubItems(5).Text = "OK"
vBangunan.SelectedItem.ListSubItems(11).Text = tDok.Text
tNIR.Text = 0
vBangunan.SetFocus

Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Description, vbCritical, "Tetnong - " & Err.Number

End Sub

Private Sub cmdClear_Click()
bersih
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPersen_Click()
On Error GoTo Salah
Dim JUMP
For i = 1 To vBangunan.ListItems.Count
    JUMP = (tPersen.Text * vBangunan.ListItems.Item(i).ListSubItems(3).Text) / 100
    If vBangunan.ListItems.Item(i).ListSubItems(5).Text = "-" Then
        vBangunan.ListItems.Item(i).ListSubItems(4).Text = Format(vBangunan.ListItems.Item(i).ListSubItems(3).Text + JUMP, "#,#0.00")
    End If
   'CALL_KELAS
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
    Pesan = "Seluruh record pada Kelurahan Terpilih akan terhapus. Lanjutkan? "
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
Me.Top = (frmUtama.ScaleHeight - Me.Height) / 2
    Me.Left = (frmUtama.Width - Me.Width) / 2
ccTahun.Clear
ccTahun.Text = Format(Now, "yyyy")
For i = Format(Now, "yyyy") To 1900 Step -1
    ccTahun.AddItem i
Next
xxTahun = (ccTahun.Text * 1) - 1
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
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
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
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_ZNT()
On Error GoTo Salah
ccZona.Clear
If ccProses = 1 Then
    QSTR = "SELECT KD_KECAMATAN,KD_ZNT FROM DAT_PETA_ZNT WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "'  GROUP BY KD_ZNT,KD_KECAMATAN ORDER BY KD_ZNT ASC "
Else
    QSTR = "SELECT KD_KECAMATAN,KD_KELURAHAN,KD_ZNT FROM DAT_PETA_ZNT WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN= '" & Left(Trim(ccKel.Text), 3) & "' GROUP BY KD_ZNT,KD_KECAMATAN,KD_KELURAHAN ORDER BY KD_ZNT ASC "
End If
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        Do While Not rPajak.EOF
            'ccZona.AddItem rPajak!KD_BLOK & "-" & rPajak!KD_ZNT
            ccZona.AddItem rPajak!KD_ZNT
        rPajak.MoveNext
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
End Sub
Sub CALL_NIR()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
vBangunan.ListItems.Clear
'QSTR = "SELECT * FROM DAT_NIR WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN= '" & Left(Trim(ccKel.Text), 3) & "' AND KD_ZNT= '" & Trim(ccZona.Text) & "' AND THN_NIR_ZNT ='" & Trim(ccTahun.Text) - 1 & "' ORDER BY KD_ZNT ASC"
If ccProses = 1 Then
    QSTR = "SELECT * FROM DAT_NIR WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "' AND THN_NIR_ZNT ='" & xxTahun & "' ORDER BY KD_ZNT ASC"
Else
    QSTR = "SELECT * FROM DAT_NIR WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN= '" & Left(Trim(ccKel.Text), 3) & "' AND THN_NIR_ZNT ='" & xxTahun & "' ORDER BY KD_ZNT ASC"
End If
'QSTR = "SELECT * FROM DAT_PETA_BLOK WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' ORDER BY KD_KELURAHAN ASC"
openDB (QSTR)
    
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
    
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_ZNT])
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Format(Trim(rPajak![NIR]), "#,#0.00")
        vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", 0
        vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 11, "", rPajak!NO_DOKUMEN
        vBangunan.ListItems.Item(i).ListSubItems.Add 12, "", rPajak!KD_KELURAHAN
        rPajak.MoveNext
    
        Loop
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
        
Screen.MousePointer = vbDefault
End Sub
Sub CALL_NIR1()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
vBangunan.ListItems.Clear
If ccProses = 1 Then
    QSTR = "SELECT * FROM DAT_NIR WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "' AND KD_ZNT= '" & Trim(ccZona.Text) & "' AND THN_NIR_ZNT ='" & xxTahun & "' ORDER BY KD_ZNT ASC"
Else
    QSTR = "SELECT * FROM DAT_NIR WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN= '" & Left(Trim(ccKel.Text), 3) & "' AND KD_ZNT= '" & Trim(ccZona.Text) & "' AND THN_NIR_ZNT ='" & xxTahun & "' ORDER BY KD_ZNT ASC"
End If
'QSTR = "SELECT * FROM DAT_NIR WHERE KD_KECAMATAN = '" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN= '" & Left(Trim(ccKel.Text), 3) & "' AND THN_NIR_ZNT ='" & Trim(ccTahun.Text) - 1 & "' ORDER BY KD_ZNT ASC"
'QSTR = "SELECT * FROM DAT_PETA_BLOK WHERE KD_KECAMATAN='" & Left(ccKec.Text, 3) & "' AND KD_KELURAHAN='" & Left(ccKel.Text, 3) & "' ORDER BY KD_KELURAHAN ASC"
openDB (QSTR)
 If rPajak.RecordCount > 0 Then rPajak.MoveFirst
        i = 0
        Do While Not rPajak.EOF
        i = i + 1
        vBangunan.ListItems.Add i, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 1, "", Format(i, "#")
        vBangunan.ListItems.Item(i).ListSubItems.Add 2, "", Trim(rPajak![KD_ZNT])
        vBangunan.ListItems.Item(i).ListSubItems.Add 3, "", Format(Trim(rPajak![NIR]), "#,#0.00")
        vBangunan.ListItems.Item(i).ListSubItems.Add 4, "", 0
        vBangunan.ListItems.Item(i).ListSubItems.Add 5, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 6, "", "-"
          vBangunan.ListItems.Item(i).ListSubItems.Add 7, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 8, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 9, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 10, "", "-"
        vBangunan.ListItems.Item(i).ListSubItems.Add 11, "", rPajak!NO_DOKUMEN
        vBangunan.ListItems.Item(i).ListSubItems.Add 12, "", rPajak!KD_KELURAHAN
        rPajak.MoveNext

        Loop
       
        
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
Salah:
If Err.Number = 0 Then Screen.MousePointer = vbDefault: Exit Sub
MsgBox Err.Number & ": " & Err.Description
Screen.MousePointer = vbDefault
End Sub

Private Sub tDok_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    ccZona.SetFocus
End If

End Sub

Private Sub tDok_LostFocus()
tDok.Text = Rep(tDok.Text)
End Sub

Private Sub tNIR_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
    cmdCari_Click
End If
End Sub

Private Sub tNIR_KeyPress(KeyAscii As Integer)
On Error Resume Next
If InStr("0123456789.,", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
End Sub

Private Sub tPersen_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
    cmdPersen_Click
End If
End Sub

Private Sub tPersen_KeyPress(KeyAscii As Integer)
On Error Resume Next

If InStr("0123456789.,-", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If

End Sub

Private Sub vBangunan_Click()
 On Error GoTo Salah
 For i = 1 To vBangunan.ListItems.Count
        vBangunan.ListItems.Item(i).ListSubItems(6).Text = "-"
    Next
 vBangunan.SelectedItem.ListSubItems(6).Text = "Proses"
    tDok.Text = vBangunan.SelectedItem.ListSubItems(11).Text
    If vBangunan.SelectedItem.ListSubItems(4).Text = "" Or vBangunan.SelectedItem.ListSubItems(4).Text = "-" Or vBangunan.SelectedItem.ListSubItems(4).Text = 0 Then
        tNIR.Text = vBangunan.SelectedItem.ListSubItems(3).Text
    Else
        tNIR.Text = vBangunan.SelectedItem.ListSubItems(4).Text
    End If
    
    For i = 1 To vBangunan.ListItems.Count
            If vBangunan.ListItems.Item(i).ListSubItems(5).Text = "OK" Then
                vBangunan.ListItems.Item(i).ListSubItems(5).Text = "OK"
            Else
                vBangunan.ListItems.Item(i).ListSubItems(5).Text = "-"
            End If
            
    Next
Salah:
If Err.Number = 0 Then Exit Sub
MsgBox Err.Description, vbCritical, "Tetnong - " & Err.Number

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
Sub CALL_KELAS()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
QSTR = "SELECT * FROM KELAS_TANAH WHERE THN_AWAL_KLS_TANAH='2011'"
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst

Do While Not rPajak.EOF
    For J = 1 To vBangunan.ListItems.Count
    If vBangunan.ListItems.Item(J).ListSubItems(4).Text * 1 >= rPajak!NILAI_MIN_TANAH * 1 And vBangunan.ListItems.Item(J).ListSubItems(4).Text * 1 <= rPajak!NILAI_MAX_TANAH * 1 Then
        vBangunan.ListItems.Item(J).ListSubItems(7).Text = Format(rPajak!KD_KLS_TANAH)
        vBangunan.ListItems.Item(J).ListSubItems(8).Text = Format(rPajak!NILAI_PER_M2_TANAH * 1000, "#,#0.00")
        vBangunan.ListItems.Item(J).ListSubItems(9).Text = Format(rPajak!NILAI_MIN_TANAH * 1000, "#,#0.00")
        vBangunan.ListItems.Item(J).ListSubItems(10).Text = Format(rPajak!NILAI_MAX_TANAH * 1000, "#,#0.00")
        
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
Sub call_SIMPAN()
On Error GoTo Salah
Screen.MousePointer = vbHourglass
If ccProses = 1 Then
    QSTR = "SELECT * FROM DAT_NIR WHERE THN_NIR_ZNT='" & ccTahun.Text & "' AND KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' "
Else
    QSTR = "SELECT * FROM DAT_NIR WHERE THN_NIR_ZNT='" & ccTahun.Text & "' AND KD_KECAMATAN='" & Left(Trim(ccKec.Text), 3) & "' AND KD_KELURAHAN='" & Left(Trim(ccKel.Text), 3) & "'"
End If
openDB (QSTR)
If rPajak.RecordCount > 0 Then rPajak.MoveFirst
If Not rPajak.EOF Then ' JIKA ADA
    If cmdSave.Caption = "&Save" Then
    MsgBox "Nilai Indikasi Rata-Rata (NIR) Tahun Pajak : " & ccTahun.Text & _
            vbCrLf & "Untuk Kelurahan : " & ccKel.Text & _
            vbCrLf & "Sudah dibuat sebelumnya...", vbCritical, "Data Exist"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
End If
If chPajak(2).Value = 1 Then GoTo Loncat1
For J = 1 To vBangunan.ListItems.Count
    If vBangunan.ListItems.Item(J).ListSubItems(4).Text = "" Or vBangunan.ListItems.Item(J).ListSubItems(4).Text = "-" Or vBangunan.ListItems.Item(J).ListSubItems(4).Text = 0 Then
        MsgBox "DATA ZNT BELUM LENGKAP/NIR MASIH KOSONG, " & _
            vbCrLf & "SILAHKAN DILENGKAPI DAHULU...", vbCritical, "DATA EMPTY"
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
Next
Loncat1:
If (cmdSave.Caption = "&Update" Or chPajak(3).Value = 1) Or (cmdSave.Caption = "&Delete" Or chPajak(2).Value = 1) Then
    Do While Not rPajak.EOF
        rPajak.Delete adAffectCurrent
        rPajak.Update
        rPajak.MoveNext
    Loop
    If cmdSave.Caption = "&Delete" Or chPajak(2).Value = 1 Then
        GoTo Keluar
    End If
End If
For i = 1 To vBangunan.ListItems.Count
'MsgBox vBangunan.ListItems.Item(i).ListSubItems(12).Text
    xxKec = Left(Trim(ccKec.Text), 3)
    xxKel = vBangunan.ListItems.Item(i).ListSubItems(12).Text
    xxZNT = vBangunan.ListItems.Item(i).ListSubItems(2).Text
    xxDok = vBangunan.ListItems.Item(i).ListSubItems(11).Text
    xxNIR = vBangunan.ListItems.Item(i).ListSubItems(4).Text
        rPajak.AddNew
        rPajak!KD_PROPINSI = "12"
        rPajak!KD_DATI2 = "12"
        rPajak!KD_ZNT = xxZNT
        rPajak!KD_KECAMATAN = xxKec
        rPajak!NIR = xxNIR
        rPajak!NO_DOKUMEN = xxDok
        rPajak!KD_KELURAHAN = xxKel
        rPajak!THN_NIR_ZNT = ccTahun.Text
        rPajak!KD_KANWIL = "01"
        rPajak!KD_KPPBB = "-"
        rPajak!JNS_DOKUMEN = "1"
    rPajak.Update
    
   'c_str = "INSERT INTO DAT_NIR(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_ZNT,THN_NIR_ZNT,KD_KANWIL,KD_KPPBB,JNS_DOKUMEN,NO_DOKUMEN,NIR)" & _
        " VALUES ('12','12','" & xxKec & "','" & xxKel & "','" & xxZNT & "','" & ccTahun & "','01','-','1','" & xxDok & "','" & xxNIR & "')"
    'openDB (c_str)
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

Sub bersih()
On Error Resume Next
ccKec.Text = ""
ccKel.Text = ""
tDok.Text = ""
tNIR.Text = 0
ccZona.Text = ""
tPersen.Text = 0
vBangunan.ListItems.Clear
End Sub
